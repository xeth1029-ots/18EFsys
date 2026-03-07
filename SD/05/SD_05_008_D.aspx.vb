Partial Class SD_05_008_D
    Inherits AuthBasePage

    Sub sUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("STUD_RESULTSTUDDATA,STUD_RESULTTWELVEDATA", objconn)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "STDNAME", Name)
        Call TIMS.sUtl_SetMaxLen(dt, "STDPID", IDNO)
        Call TIMS.sUtl_SetMaxLen(dt, "Q8OTHER", Q8Other)
    End Sub

    '結訓學員資料卡登錄

    'SD_11_001.aspx
    'STUD_DATALID
    'UPDATE TABLE:  STUD_RESULTSTUDDATA
    'update table:  Stud_ResultIdentData
    'update table:  Stud_ResultTwelveData

    Const cst_search As String = "search"

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
        Call sUtl_PageInit1()
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Me.TODAY1YEAR.Value = Year(Now)
        '判斷是否有OCID有OCID為署(局)屬
        HidOCID.Value = TIMS.ClearSQM(Request("OCID"))
        If HidOCID.Value <> "" Then '署(局)屬
            HidSTDate.Value = TIMS.GET_OCIDInfo(HidOCID.Value, "SDate", objconn)
            Me.TODAY1YEAR.Value = Year(HidSTDate.Value)
        End If

        If Not IsPostBack Then
            '增加鍵值
            Call Add_Items()
            Call Cretae1()

            Button1.Attributes("onclick") = "javascript:return chkdata();"
            Button2.Attributes("onclick") = "javascript:return chkdata();"
            Q9a.Attributes("onclick") = "change()"
            Q9b.Attributes("onclick") = "change()"
            'Q11a.Attributes("onclick") = "change()"
            'Q11b.Attributes("onclick") = "change()"
            Q8.Attributes("onchange") = "change()"
            'IdentityID.Attributes("onclick") = "GetIdent();"

            IDNO.Attributes("onchange") = "javascript:chkidnosex();"
            IDNO.Attributes("onblur") = "javascript:chkidnosex();"
        End If

    End Sub

    Sub Cretae1()
        '"01,04,05,06,07,09,10,11,13,14,26,28,30,31,32,33,37,40,43,44"
        'Dim blnRedirectFlag As Boolean = False
        Button1.Visible = False
        Button2.Visible = False
        Table3.Visible = False
        Table4.Style.Item("display") = "none"

        Dim sql As String = ""
        Dim rqProecess As String = TIMS.ClearSQM(Request("Proecess"))
        Dim rqDLID As String = TIMS.ClearSQM(Request("DLID"))
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))

        Select Case rqProecess
            Case "addmyall" '署(局)屬 新增-有封面,導向所有學生建立 addmyall
                '檢查是否有封面
                sql = "SELECT * FROM STUD_DATALID WHERE DLID='" & rqDLID & "'"
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                If dr Is Nothing Then
                    TIMS.Utl_Redirect1(Me, "SD_05_008_C.aspx?Proecess=addmy&OCID=" & rqOCID)
                    Return 'Exit Sub
                End If
                DLID.Value = Convert.ToString(dr("DLID"))
                Table3.Visible = True
                Call Add_Student()
                Call dis_data()

            Case "add" '【署(局)屬】
                '檢查是否有封面
                sql = "SELECT * FROM STUD_DATALID WHERE DLID='" & rqDLID & "' AND OCID IS NOT NULL"
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                If dr Is Nothing Then
                    '查無封面
                    TIMS.Utl_Redirect1(Me, "SD_05_008_C.aspx?Proecess=addmy&OCID=" & rqOCID)
                    Return
                End If
                DLID.Value = Convert.ToString(dr("DLID"))
                Button1.Visible = False
                Button2.Visible = True
                Table4.Style.Item("display") = "" '學員明細

                Dim rqSOCID As String = TIMS.ClearSQM(Request("SOCID"))
                Call create_basic(rqSOCID) '單筆
            Case "addother"
                '非署(局)屬。'檢查是否有封面
                sql = "SELECT * FROM STUD_DATALID WHERE DLID='" & rqDLID & "' AND OCID IS NULL"
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                If dr Is Nothing Then
                    '查無封面
                    TIMS.Utl_Redirect1(Me, "SD_05_008_C.aspx?Proecess=addother&DLID=" & rqDLID)
                    Return 'Exit Sub
                End If
                DLID.Value = Convert.ToString(dr("DLID"))
                Dim rstA2 As Boolean = Actoin_cnt2(rqDLID)
                Dim s_MRqID As String = TIMS.Get_MRqID(Me)
                If rstA2 Then
                    Page.RegisterStartupScript(TIMS.xBlockName, "<script>location.href='SD_05_008.aspx?ID=" & s_MRqID & "';</script>")
                    Return ' Exit Sub
                End If
                'dr("ResultCount")
                'Dim iResultCount As Integer = 0 'Val(dr("ResultCount"))
                'Dim iStdCount As Integer = 0 '
                'If Convert.ToString(dr("ResultCount")) <> "" Then iResultCount = CInt(dr("ResultCount"))
                'sql = "SELECT COUNT(1) cnt FROM STUD_RESULTSTUDDATA WHERE DLID='" & rqDLID & "'"
                'iStdCount = DbAccess.ExecuteScalar(sql, objconn)
                'If iStdCount >= iResultCount Then
                '    Common.MessageBox(Me, "已超過結訓人數，不得新增!")
                '    blnRedirectFlag = True
                '    Page.RegisterStartupScript(TIMS.xBlockName, "<script>location.href='SD_05_008.aspx?ID=" & Request("ID") & "';</script>")
                '    Exit Sub
                'End If

                Button1.Visible = True
                Button2.Visible = True
                Table4.Style.Item("display") = "" '學員明細
            Case "edit" '顯示【非署(局)屬】
                Button1.Visible = False
                Button2.Visible = True
                Table4.Style.Item("display") = "" '學員明細
                '顯示【非署(局)屬】
                Call create_basic2()
        End Select

        'If Not Session(cst_search) Is Nothing And Not blnRedirectFlag Then
        If Session(cst_search) IsNot Nothing Then ViewState(cst_search) = Session(cst_search)

    End Sub

    '控制所有欄位讀取。
    Sub dis_data()
        'Name.ReadOnly = True
        'StudentID.ReadOnly = True
        'IDNO.ReadOnly = True

        Name.Attributes.Add("onkeydown", "this.blur()")
        Name.Attributes.Add("oncontextmenu", "return false;")
        StudentID.Attributes.Add("onkeydown", "this.blur()")
        StudentID.Attributes.Add("oncontextmenu", "return false;")
        IDNO.Attributes.Add("onkeydown", "this.blur()")
        IDNO.Attributes.Add("oncontextmenu", "return false;")
        Sex.Enabled = False

        'Byear.ReadOnly = True
        'Bmonth.ReadOnly = True
        'Bday.ReadOnly = True
        Byear.Attributes.Add("onkeydown", "this.blur()")
        Byear.Attributes.Add("oncontextmenu", "return false;")
        Bmonth.Attributes.Add("onkeydown", "this.blur()")
        Bmonth.Attributes.Add("oncontextmenu", "return false;")
        Bday.Attributes.Add("onkeydown", "this.blur()")
        Bday.Attributes.Add("oncontextmenu", "return false;")

        DegreeID.Enabled = False
        MilitaryID.Enabled = False

        IdentityID.Enabled = False
        IdentityValue.Value = ""
        For Each item As ListItem In IdentityID.Items
            If item.Selected AndAlso item.Value <> "" Then
                IdentityValue.Value &= String.Concat(If(IdentityValue.Value <> "", ",", ""), item.Value)
            End If
        Next
    End Sub

    Sub Add_Student()
        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))

        Dim sParms As New Hashtable
        sParms.Add("OCID", rqOCID)
        Dim sSql As String = ""
        sSql &= " SELECT a.Name,a.SOCID,a.StudentID,c.DLID,a.StudStatus" & vbCrLf
        sSql &= " FROM V_STUDENTINFO a" & vbCrLf
        sSql &= " LEFT JOIN STUD_RESULTSTUDDATA c ON c.SOCID=a.SOCID" & vbCrLf
        'sSql &= " WHERE a.TPLANID ='06' and a.YEARS='2022'" & vbCrLf
        sSql &= " WHERE a.OCID=@OCID"
        sSql &= " AND a.StudStatus NOT IN (2,3)" '排除離退訓學員
        sSql &= " ORDER BY a.StudentID"
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)

        Dim n1flag As Boolean = False '無未填寫學員
        SOCID.Items.Clear()
        SOCID.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, 0))
        For Each dr As DataRow In dt.Rows
            If Convert.ToString(dr("DLID")) = "" Then
                n1flag = True '有未填寫學員
                SOCID.Items.Add(New ListItem(String.Concat(dr("Name"), "(", Right(dr("StudentID"), 2), ")"), dr("SOCID")))
            End If
        Next
        '無未填寫學員
        If Not n1flag Then
            Page.RegisterStartupScript(TIMS.xBlockName, "<script>alert('所有學員皆填寫完畢');location.href='SD_05_008.aspx?ID=" & s_MRqID & "';</script>")
        End If
    End Sub

    '常用鍵詞
    Sub Add_Items()
        Dim sql As String = ""
        '學歷鍵值
        sql = "SELECT DEGREEID,NAME FROM KEY_DEGREE WHERE DEGREETYPE=1 ORDER BY SORT"
        DegreeID.Items.Clear()
        DbAccess.MakeListItem(DegreeID, sql, objconn)
        DegreeID.Items.Insert(0, New ListItem("==請選擇==", 0))

        '兵役鍵值 
        sql = "SELECT MILITARYID,NAME FROM KEY_MILITARY WHERE MILITARYID !='00' ORDER BY MilitaryID"
        MilitaryID.Items.Clear()
        DbAccess.MakeListItem(MilitaryID, sql, objconn)
        MilitaryID.Items.Insert(0, New ListItem("==請選擇==", "00"))

        'sPrtFN1 = "ResultStud" '署(局)屬
        Dim blnPrint2016 As Boolean = False '【非署(局)屬】或 舊【署(局)屬】
        If HidOCID.Value <> "" Then '署(局)屬
            If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuySD20160701), CDate(HidSTDate.Value)) >= 0 Then
                '身分別鍵值IdentityID '201607 'blnPrint2016 = True 'sPrtFN1 = "ResultStud10507" '署(局)屬
                blnPrint2016 = True '署(局)屬
            End If
        End If
        'If TIMS.sUtl_ChkTest() Then blnPrint2016 = True '測試201607
        'If blnPrint2016 Then
        '    '身分別鍵值IdentityID
        '    IdentityID = TIMS.Get_Identity(IdentityID, 11, objconn)
        'Else
        '    '【非署(局)屬】 或 舊【署(局)屬】 '身分別鍵值IdentityID
        '    IdentityID = TIMS.Get_Identity(IdentityID, 1, objconn)
        'End If
        IdentityID = TIMS.Get_Identity(IdentityID, 53, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)

        Dim rqJuzhu As String = TIMS.ClearSQM(Request("Juzhu"))
        'Dim Juzhu As String = Request("Juzhu")
        trq12y2013.Visible = False
        trq12y2014.Visible = False
        'Select Case Juzhu
        '    Case "1" '署(局)屬
        '        trq12y2014.Visible = True
        '    Case "2" '非署(局)屬
        '        trq12y2013.Visible = True
        'End Select
        'trq12y2014 '署(局)屬  '非署(局)屬
        'trq12y2013.Visible = False
        trq12y2014.Visible = True

    End Sub

    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        Table4.Style.Item("display") = ""
        Button1.Visible = True
        Button2.Visible = True
        Call create_basic(SOCID.SelectedValue)          '尋找出個人的基本資料
    End Sub

    '清除所有資料欄位。
    Sub Clean_Data()
        SOCIDNum.Value = ""
        Name.Text = ""
        StudentID.Text = ""
        IDNO.Text = ""
        If Not Sex.SelectedItem Is Nothing Then
            Sex.SelectedItem.Selected = False
        End If
        Byear.Text = ""
        Bmonth.Text = ""
        Bday.Text = ""
        If Not DegreeID.SelectedItem Is Nothing Then
            DegreeID.SelectedItem.Selected = False
        End If
        If Not MilitaryID.SelectedItem Is Nothing Then
            MilitaryID.SelectedItem.Selected = False
        End If
        For i As Integer = 0 To IdentityID.Items.Count - 1
            IdentityID.Items(i).Selected = False
        Next
        If Not Q7.SelectedItem Is Nothing Then
            Q7.SelectedItem.Selected = False
        End If
        If Not Q8.SelectedItem Is Nothing Then
            Q8.SelectedItem.Selected = False
        End If
        Q8Other.Text = ""
        Q9a.Checked = False
        Q9b.Checked = False
        For i As Integer = 0 To Q9Y.Items.Count - 1
            Q9Y.Items(i).Selected = False
        Next
        If Not Q10.SelectedItem Is Nothing Then
            Q10.SelectedItem.Selected = False
        End If
        'Q11a.Checked = False
        'Q11b.Checked = False
        If Not Q11N.SelectedItem Is Nothing Then
            Q11N.SelectedItem.Selected = False
        End If

        'trq12y2013
        Q12AY.Checked = False
        Q12AN.Checked = False
        For i As Integer = 0 To Q12B.Items.Count - 1
            Q12B.Items(i).Selected = False
        Next
        'trq12y2014
        Q12V1.SelectedIndex = -1
        Q12V2.SelectedIndex = -1
        Q12V3.SelectedIndex = -1
        Q12V4.SelectedIndex = -1
        Q12V5.SelectedIndex = -1

        'For i = 0 To Q12.Items.Count - 1
        '    Q12.Items(i).Selected = False
        'Next
        'Q12Other.Text = ""
    End Sub

    '學員基本資料的顯示。【署(局)屬】 Stud_StudentInfo Class_StudentsOfClass
    Sub create_basic(ByVal SOCIDValue As String)
        'If Me.ViewState("data") = 1 Then            '表示資料有被呼叫過，要先清除
        Call Clean_Data()
        'End If
        'Me.ViewState("data") = 1
        SOCIDValue = TIMS.ClearSQM(SOCIDValue)
        SOCIDNum.Value = SOCIDValue

        Dim sPMS As New Hashtable
        sPMS.Add("SOCID", SOCIDValue)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.Name" & vbCrLf
        sql &= " ,a.Sex" & vbCrLf
        sql &= " ,a.Birthday" & vbCrLf
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " ,a.DegreeID" & vbCrLf
        sql &= " ,a.MilitaryID" & vbCrLf
        sql &= " ,b.IdentityID" & vbCrLf
        sql &= " ,b.SOCID" & vbCrLf
        sql &= " ,b.StudentID" & vbCrLf
        sql &= " ,c.DLID" & vbCrLf
        sql &= " FROM STUD_STUDENTINFO a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON b.SID=a.SID" & vbCrLf
        sql &= " LEFT JOIN STUD_RESULTSTUDDATA c ON c.SOCID=b.SOCID" & vbCrLf
        sql &= " WHERE b.SOCID=@SOCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, sPMS)

        'Dim dr As DataRow = Nothing
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)

            Name.Text = Convert.ToString(dr("Name"))
            StudentID.Text = Right(Convert.ToString(dr("StudentID")), 2)
            IDNO.Text = Convert.ToString(dr("IDNO"))
            If Convert.ToString(dr("Sex")) = "M" Then
                Sex.Items(0).Selected = True
            ElseIf Convert.ToString(dr("Sex")) = "F" Then
                Sex.Items(1).Selected = True
            End If
            Dim all() As String
            If Convert.ToString(dr("Birthday")) <> "" Then
                all = Split(dr("Birthday"), "/", , CompareMethod.Text)
                Byear.Text = all(0)
                Bmonth.Text = all(1)
                Bday.Text = all(2)
            End If
            'Dim i, j
            For i As Integer = 0 To DegreeID.Items.Count - 1
                If DegreeID.Items(i).Value = Convert.ToString(dr("DegreeID")) Then
                    DegreeID.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To MilitaryID.Items.Count - 1
                If MilitaryID.Items(i).Value = Convert.ToString(dr("MilitaryID")) Then
                    MilitaryID.Items(i).Selected = True
                End If
            Next

            Call TIMS.SetCblValue(IdentityID, Convert.ToString(dr("IdentityID")))
            'If Convert.ToString(dr("IdentityID")) <> "" Then
            '    all = Split(dr("IdentityID"), ",")
            '    For i As Integer = 0 To IdentityID.Items.Count - 1
            '        For j As Integer = 0 To all.Length - 1
            '            If all(j) = IdentityID.Items(i).Value Then
            '                IdentityID.Items(i).Selected = True
            '            End If
            '        Next
            '    Next
            'End If
        End If

        Call dis_data()

    End Sub

    '學員結訓學員資料卡的顯示。【非署(局)屬】 STUD_RESULTSTUDDATA
    Sub create_basic2()
        SOCIDNum.Value = ""
        Call Clean_Data()
        Dim rqDLID As String = TIMS.ClearSQM(Request("DLID"))
        Dim rqSubNo As String = TIMS.ClearSQM(Request("SubNo"))

        'Dim i As Integer = 0
        'Dim j As Integer = 0
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.DLID" & vbCrLf '/*PK*/
        sql &= " ,a.SUBNO" & vbCrLf '/*PK*/
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.STDNAME" & vbCrLf
        sql &= " ,a.STUDENTID" & vbCrLf
        sql &= " ,a.STDPID" & vbCrLf
        sql &= " ,a.SEX" & vbCrLf
        sql &= " ,a.BIRTHYEAR" & vbCrLf
        sql &= " ,a.BIRTHMONTH" & vbCrLf
        sql &= " ,a.BIRTHDATE" & vbCrLf
        sql &= " ,a.DEGREEID" & vbCrLf
        sql &= " ,a.MILITARYID" & vbCrLf
        sql &= " ,a.Q7" & vbCrLf 'a.Q7
        sql &= " ,a.Q8" & vbCrLf
        sql &= " ,a.Q8OTHER" & vbCrLf
        sql &= " ,a.Q9" & vbCrLf
        sql &= " ,a.Q9Y" & vbCrLf
        sql &= " ,a.Q10" & vbCrLf
        sql &= " ,a.Q11" & vbCrLf
        sql &= " ,a.Q11N" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.Q12A" & vbCrLf
        sql &= " ,a.Q12B" & vbCrLf
        sql &= " ,a.Q12V1" & vbCrLf
        sql &= " ,a.Q12V2" & vbCrLf
        sql &= " ,a.Q12V3" & vbCrLf
        sql &= " ,a.Q12V4" & vbCrLf
        sql &= " ,a.Q12V5" & vbCrLf
        sql &= " FROM STUD_RESULTSTUDDATA a" & vbCrLf
        sql &= " WHERE a.DLID='" & rqDLID & "' AND a.SUBNO='" & rqSubNo & "'"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)

        'Dim i, j
        If Not dr Is Nothing Then
            DLID.Value = Convert.ToString(dr("DLID"))
            SubNo.Value = Convert.ToString(dr("SubNo"))
            Name.Text = Convert.ToString(dr("StdName"))
            SOCIDNum.Value = Convert.ToString(dr("SOCID"))
            StudentID.Text = Convert.ToString(dr("StudentID"))
            IDNO.Text = Convert.ToString(dr("StdPID"))
            If Convert.ToString(dr("Sex")) = "1" Then
                Sex.Items(0).Selected = True
            ElseIf Convert.ToString(dr("Sex")) = "2" Then
                Sex.Items(1).Selected = True
            End If
            Byear.Text = Convert.ToString(dr("BirthYear"))
            Bmonth.Text = Convert.ToString(dr("BirthMonth"))
            Bday.Text = Convert.ToString(dr("BirthDate"))
            'Dim all()

            For i As Integer = 0 To DegreeID.Items.Count - 1
                If DegreeID.Items(i).Value = Convert.ToString(dr("DegreeID")) Then
                    DegreeID.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To MilitaryID.Items.Count - 1
                If MilitaryID.Items(i).Value = Convert.ToString(dr("MilitaryID")) Then
                    MilitaryID.Items(i).Selected = True
                End If
            Next


            For i As Integer = 0 To Q7.Items.Count - 1
                If Q7.Items(i).Value = Convert.ToString(dr("Q7")) Then
                    Q7.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To Q8.Items.Count - 1
                If Q8.Items(i).Value = Convert.ToString(dr("Q8")) Then
                    Q8.Items(i).Selected = True
                End If
            Next
            Q8Other.Text = Convert.ToString(dr("Q8Other"))
            If Convert.ToString(dr("Q9")) = "Y" Then
                Q9a.Checked = True
            ElseIf Convert.ToString(dr("Q9")) = "N" Then
                Q9b.Checked = True
            End If
            For i As Integer = 0 To Q9Y.Items.Count - 1
                If Q9Y.Items(i).Value = Convert.ToString(dr("Q9Y")) Then
                    Q9Y.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To Q10.Items.Count - 1
                If Q10.Items(i).Value = Convert.ToString(dr("Q10")) Then
                    Q10.Items(i).Selected = True
                End If
            Next
            'If Convert.ToString(dr("Q11")) = "Y" Then
            '    Q11a.Checked = True
            'End If
            'If Convert.ToString(dr("Q11")) = "N" Then
            '    Q11b.Checked = True
            'End If
            If Convert.ToString(dr("Q11N")) <> "" Then
                For i As Integer = 0 To Q11N.Items.Count - 1
                    If Q11N.Items(i).Value = Convert.ToString(dr("Q11N")) Then
                        Q11N.Items(i).Selected = True
                    End If
                Next
            End If

            'trq12y2013
            If Convert.ToString(dr("Q12A")) = "Y" Then
                Q12AY.Checked = True
            ElseIf Convert.ToString(dr("Q12A")) = "N" Then
                Q12AN.Checked = True
                Call TIMS.SetCblValue(Q12B, Convert.ToString(dr("Q12B")))
            End If

            'trq12y2014
            If Convert.ToString(dr("Q12V1")) <> "" Then Common.SetListItem(Q12V1, Convert.ToString(dr("Q12V1")))
            If Convert.ToString(dr("Q12V2")) <> "" Then Common.SetListItem(Q12V2, Convert.ToString(dr("Q12V2")))
            If Convert.ToString(dr("Q12V3")) <> "" Then Common.SetListItem(Q12V3, Convert.ToString(dr("Q12V3")))
            If Convert.ToString(dr("Q12V4")) <> "" Then Common.SetListItem(Q12V4, Convert.ToString(dr("Q12V4")))
            If Convert.ToString(dr("Q12V5")) <> "" Then Common.SetListItem(Q12V5, Convert.ToString(dr("Q12V5")))
        End If

        'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
        'BY AMU 2009-07-30
        ''參訓身分
        'sql = "SELECT * FROM Stud_ResultIdentData WHERE DLID='" & Request("DLID") & "' and SubNo='" & Request("SubNo") & "'"
        'Dim dt As DataTable = DbAccess.GetDataTable(sql)
        'For i = 0 To dt.Rows.Count - 1
        '    For j = 0 To IdentityID.Items.Count - 1
        '        If dt.Rows(i).Item("IdentityID") = IdentityID.Items(j).Value Then
        '            IdentityID.Items(j).Selected = True
        '        End If
        '    Next
        'Next

        'For j As Integer = 0 To IdentityID.Items.Count - 1
        '    IdentityID.Items(j).Selected = False
        'Next
        Call TIMS.SetCblValue(IdentityID, "")

        'Dim dt As DataTable = Nothing
        ''參訓身分【非署(局)屬查詢署(局)屬可能性資料。】有資料依
        sql = "" & vbCrLf
        sql &= " SELECT cs.IdentityID " & vbCrLf
        sql &= " FROM STUD_RESULTSTUDDATA sr" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on sr.SOCID=cs.SOCID" & vbCrLf
        sql &= " WHERE sr.DLID='" & rqDLID & "' AND sr.SubNo='" & rqSubNo & "'" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            '署(局)屬狀況 【JOIN 到 Class_StudentsOfClass】
            'Dim arIden As String() = dr("IdentityID").ToString.Split(",")
            'For i As Integer = 0 To arIden.Length - 1
            'Next
            Dim tmpIdentityID As String = Convert.ToString(dr("IdentityID"))
            Call TIMS.SetCblValue(IdentityID, tmpIdentityID)
            'If tmpIdentityID <> "" Then
            '    For j As Integer = 0 To IdentityID.Items.Count - 1
            '        'IdentityID.Items(j).Selected = False
            '        If IdentityID.Items(j).Value <> "" _
            '            AndAlso tmpIdentityID.IndexOf(IdentityID.Items(j).Value) > -1 Then
            '            IdentityID.Items(j).Selected = True
            '        End If
            '    Next
            'End If
            Call dis_data() '取消輸入項

        Else
            '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
            sql = "SELECT * FROM STUD_RESULTIDENTDATA WHERE DLID='" & rqDLID & "' and SubNo='" & rqSubNo & "'"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
            For Each dr1 As DataRow In dt.Rows
                For j As Integer = 0 To IdentityID.Items.Count - 1
                    'IdentityID.Items(j).Selected = False
                    If Convert.ToString(dr1("IdentityID")) = IdentityID.Items(j).Value Then
                        IdentityID.Items(j).Selected = True
                    End If
                Next
            Next

        End If

        '判斷是否要隱藏
        sql = "SELECT * FROM STUD_DATALID WHERE DLID='" & rqDLID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then
            Select Case dr("UnitCode").ToString '參考統計室-分署代碼表
                Case "001", "002", "003", "004", "005", "006"
                    Call dis_data() '取消輸入項
            End Select
        End If

        ''第12題
        ''Dim dt As DataTable
        'sql = "SELECT * FROM Stud_ResultTwelveData WHERE DLID='" & rqDLID & "' and SubNo='" & rqSubNo & "'"
        'dt = DbAccess.GetDataTable(sql, objconn)
        'For i = 0 To dt.Rows.Count - 1
        '    For j = 0 To Q12.Items.Count - 1
        '        If Q12.Items(j).Value = Convert.ToString(dt.Rows(i).Item("Q12")) Then
        '            Q12.Items(j).Selected = True
        '        End If
        '        If Convert.ToString(dt.Rows(i).Item("Q12")) = "5" Then
        '            Q12Other.Text = dt.Rows(i).Item("Q12Other").ToString
        '        End If
        '    Next
        'Next

    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Name.Text = TIMS.ClearSQM(Name.Text)
        StudentID.Text = TIMS.ClearSQM(StudentID.Text)
        If Name.Text = "" Then
            Errmsg += "姓名資料有誤，不可為空" & vbCrLf
        Else
            If Name.Text.Length <= 1 Then
                Errmsg += "姓名資料長度有誤，不可小於等於1" & vbCrLf
            End If
        End If

        If StudentID.Text.Trim = "" Then
            Errmsg += "學號資料有誤，不可為空" & vbCrLf
        End If
        If SOCIDNum.Value.Trim <> "" Then '署(局)屬不會為空
            If Not IsNumeric(SOCIDNum.Value.Trim) Then
                'Errmsg += "局屬學號資料有誤，應為數字" & vbCrLf
                Errmsg += "署屬學號資料有誤，應為數字" & vbCrLf
            End If
        End If
        Select Case Sex.SelectedValue
            Case "1", "2"
            Case Else
                Errmsg += "性別資料有誤，不可為空" & vbCrLf
        End Select

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        SOCIDNum.Value = TIMS.ClearSQM(SOCIDNum.Value)

        If IDNO.Text <> "" Then
            If SOCIDNum.Value <> "" Then '署(局)屬才驗證
                '1:國民身分證 
                Dim flag1 As Boolean = TIMS.CheckIDNO(IDNO.Text)
                '2:居留證 4:居留證2021
                Dim flag2 As Boolean = TIMS.CheckIDNO2(IDNO.Text, 2)
                Dim flag4 As Boolean = TIMS.CheckIDNO2(IDNO.Text, 4)

                If Not flag1 AndAlso Not flag2 AndAlso Not flag4 Then
                    Errmsg += "身分證號碼(或居留證號) 格式有誤!" & vbCrLf
                End If
            End If
        Else
            Errmsg += "身分證號碼資料有誤，不可為空" & vbCrLf
        End If

        Try
            If Byear.Text.Trim <> "" Then Byear.Text = TIMS.ChangeIDNO(Byear.Text)
            If Bmonth.Text.Trim <> "" Then Bmonth.Text = TIMS.ChangeIDNO(Bmonth.Text)
            If Bday.Text.Trim <> "" Then Bday.Text = TIMS.ChangeIDNO(Bday.Text)
            Dim vDate1 As String = ""
            vDate1 = ""
            vDate1 += Convert.ToString(Byear.Text)
            vDate1 += "/" & Convert.ToString(Bmonth.Text)
            vDate1 += "/" & Convert.ToString(Bday.Text)
            If IsDate(vDate1) Then
                vDate1 = CDate(vDate1).ToString("yyyy/MM/dd")
                Byear.Text = Year(CDate(vDate1))
                Bmonth.Text = Month(CDate(vDate1))
                Bday.Text = Day(CDate(vDate1))
            Else
                Errmsg += "出生日期資料有誤，應為日期格式" & vbCrLf
            End If
        Catch ex As Exception
            Errmsg += "出生日期資料有誤，應為日期格式" & vbCrLf
        End Try

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Sub SaveData1()

    End Sub

    '儲存資料 (Button2:填寫完畢)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        Dim sql As String = ""
        Dim dt1 As DataTable = Nothing

        Dim dt2 As DataTable = Nothing
        Dim dr2 As DataRow = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        Dim sMsgbox As String = "未完成"
        'Dim Proecess As String = rqProecess
        Dim rqProecess As String = TIMS.ClearSQM(Request("Proecess"))
        Dim rqDLID As String = TIMS.ClearSQM(Request("DLID"))

        Select Case rqProecess
            Case "add", "addmyall", "addother"
                sql = ""
                sql &= " SELECT * FROM STUD_RESULTSTUDDATA"
                sql &= " WHERE DLID='" & DLID.Value & "'" & vbCrLf
                sql &= " AND StdPID='" & IDNO.Text & "'" & vbCrLf
                Dim drD As DataRow = DbAccess.GetOneRow(sql, objconn)
                If Not drD Is Nothing Then
                    Common.MessageBox(Me, "此筆學員資料已存在,無法再次儲存!!")
                    If Session(cst_search) Is Nothing Then Session(cst_search) = Me.ViewState(cst_search)

                    Page.RegisterStartupScript(TIMS.xBlockName, "<script>location.href='SD_05_008.aspx?ID=" & s_MRqID & "';</script>")
                    Exit Sub
                End If
        End Select


        Dim i_SubNo As Integer = -1
        Dim dr1 As DataRow = Nothing
        Dim da1 As SqlDataAdapter = Nothing
        Select Case rqProecess
            Case "add", "addmyall", "addother"
                i_SubNo = TIMS.Get_nSubNOxResultStudData(DLID.Value, objconn)
                sql = "SELECT * FROM STUD_RESULTSTUDDATA WHERE 1<>1"
                dt1 = DbAccess.GetDataTable(sql, da1, objconn)
                dr1 = dt1.NewRow()
                dt1.Rows.Add(dr1)
                sMsgbox = "新增"
            Case "edit"
                sql = "SELECT * FROM STUD_RESULTSTUDDATA WHERE DLID='" & DLID.Value & "' and SubNo='" & SubNo.Value & "'"
                dt1 = DbAccess.GetDataTable(sql, da1, objconn)
                dr1 = dt1.Rows(0)
                i_SubNo = Val(TIMS.ClearSQM(SubNo.Value))
                sMsgbox = "編輯"
            Case Else
                Common.MessageBox(Me, "作業狀態異常，請重新查詢!!")
                Page.RegisterStartupScript(TIMS.xBlockName, "<script>location.href='SD_05_008.aspx?ID=" & s_MRqID & "';</script>")
                Return 'Exit Sub
        End Select

        'STUD_RESULTSTUDDATA表格
        Dim V_Sex As String = TIMS.GetListValue(Sex)
        Dim V_DegreeID As String = TIMS.GetListValue(DegreeID)
        Dim V_MilitaryID As String = TIMS.GetListValue(MilitaryID)
        Dim V_Q7 As String = TIMS.GetListValue(Q7)
        Dim V_Q8 As String = TIMS.GetListValue(Q8)
        Dim V_Q9Y As String = TIMS.GetListValue(Q9Y)
        Dim V_Q10 As String = TIMS.GetListValue(Q10)
        Dim V_Q11N As String = TIMS.GetListValue(Q11N)
        Q8Other.Text = TIMS.ClearSQM(Q8Other.Text)

        SubNo.Value = Convert.ToString(i_SubNo)
        dr1("DLID") = DLID.Value
        dr1("SubNo") = i_SubNo ' SubNo.Value
        dr1("SOCID") = If(SOCIDNum.Value <> "", SOCIDNum.Value, Convert.DBNull)
        dr1("StdName") = Name.Text
        dr1("StudentID") = TIMS.ClearSQM(StudentID.Text)
        dr1("StdPID") = TIMS.ChangeIDNO(IDNO.Text)
        dr1("Sex") = If(V_Sex = "1", V_Sex, If(V_Sex = "2", V_Sex, Convert.DBNull))
        dr1("BirthYear") = Convert.ToInt32(Byear.Text)
        dr1("BirthMonth") = Convert.ToInt32(Bmonth.Text)
        dr1("BirthDate") = Convert.ToInt32(Bday.Text)
        dr1("DegreeID") = If(V_DegreeID <> "", V_DegreeID, Convert.DBNull)
        dr1("MilitaryID") = If(V_MilitaryID <> "", V_MilitaryID, Convert.DBNull)
        dr1("Q7") = If(V_Q7 <> "", V_Q7, Convert.DBNull)
        dr1("Q8") = If(V_Q8 <> "", V_Q8, Convert.DBNull)
        dr1("Q8Other") = If(Q8Other.Text <> "", Q8Other.Text, Convert.DBNull)
        dr1("Q9") = If(Q9a.Checked, "Y", If(Q9b.Checked, "N", Convert.DBNull))
        dr1("Q9Y") = If(Q9a.Checked AndAlso V_Q9Y <> "", V_Q9Y, Convert.DBNull) 'Q9Y.SelectedValue
        dr1("Q10") = If(V_Q10 <> "", V_Q10, Convert.DBNull)
        dr1("Q11") = Convert.DBNull '(Y/N)
        '非 4:留在原場(廠)服務
        dr1("Q11N") = If(V_Q11N <> "", V_Q11N, Convert.DBNull)
        'trq12y2013
        Dim Q12Bval As String = TIMS.GetCblValue(Q12B)
        If trq12y2013.Visible Then
            dr1("Q12A") = If(Q12AY.Checked, "Y", If(Q12AN.Checked, "N", Convert.DBNull))
            dr1("Q12B") = If(Q12AN.Checked AndAlso Q12Bval <> "", Q12Bval, Convert.DBNull)
        End If
        'trq12y2014
        If trq12y2014.Visible Then
            dr1("Q12V1") = If(Q12V1.SelectedValue <> "", Q12V1.SelectedValue, Convert.DBNull)
            dr1("Q12V2") = If(Q12V2.SelectedValue <> "", Q12V2.SelectedValue, Convert.DBNull)
            dr1("Q12V3") = If(Q12V3.SelectedValue <> "", Q12V3.SelectedValue, Convert.DBNull)
            dr1("Q12V4") = If(Q12V4.SelectedValue <> "", Q12V4.SelectedValue, Convert.DBNull)
            dr1("Q12V5") = If(Q12V5.SelectedValue <> "", Q12V5.SelectedValue, Convert.DBNull)
        End If
        dr1("ModifyAcct") = sm.UserInfo.UserID
        dr1("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt1, da1) 'objconn

        Dim tConn As SqlConnection = DbAccess.GetConnection
        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            'Stud_ResultIdentData表格
            sql = "DELETE STUD_RESULTIDENTDATA WHERE DLID='" & DLID.Value & "' AND SUBNO='" & SubNo.Value & "'"
            DbAccess.ExecuteNonQuery(sql, objTrans)
            sql = "SELECT * FROM STUD_RESULTIDENTDATA WHERE 1<>1"
            dt2 = DbAccess.GetDataTable(sql, da2, objTrans)
            'Stud_ResultIdentData表格
            'Dim i As Integer = 0
            If IdentityValue.Value = "" Then
                For i As Integer = 0 To IdentityID.Items.Count - 1
                    If IdentityID.Items(i).Selected Then
                        dr2 = dt2.NewRow()
                        dt2.Rows.Add(dr2)
                        dr2("DLID") = DLID.Value
                        dr2("SubNo") = SubNo.Value
                        dr2("IdentityID") = IdentityID.Items(i).Value
                    End If
                Next
            Else
                Dim MyArray As Array = Split(IdentityValue.Value, ",")
                For i As Integer = 0 To MyArray.Length - 1
                    dr2 = dt2.NewRow()
                    dt2.Rows.Add(dr2)
                    dr2("DLID") = DLID.Value
                    dr2("SubNo") = SubNo.Value
                    dr2("IdentityID") = MyArray(i)
                Next
            End If
            DbAccess.UpdateDataTable(dt2, da2, objTrans)

            'Call TIMS.OpenDbConn(tConn)
            'Stud_ResultTwelveData表格
            sql = "DELETE STUD_RESULTTWELVEDATA WHERE DLID='" & DLID.Value & "' and SubNo='" & SubNo.Value & "'"
            DbAccess.ExecuteNonQuery(sql, objTrans)

            'sql = "SELECT * FROM Stud_ResultTwelveData WHERE 1<>1"
            'dt3 = DbAccess.GetDataTable(sql, da3, objTrans)
            ''Stud_ResultTwelveData表格
            'For i = 0 To Q12.Items.Count - 1
            '    If Q12.Items(i).Selected = True Then
            '        dr3 = dt3.NewRow()
            '        dt3.Rows.Add(dr3)
            '        dr3("DLID") = DLID.Value
            '        dr3("SubNo") = SubNo.Value
            '        dr3("Q12") = Q12.Items(i).Value
            '        If Q12Other.Text = "" Then
            '            dr3("Q12Other") = Convert.DBNull
            '        Else
            '            dr3("Q12Other") = Q12Other.Text
            '        End If
            '    End If
            'Next
            'DbAccess.UpdateDataTable(dt3, da3, objTrans)
            DbAccess.CommitTrans(objTrans)

            sMsgbox += "成功"

            If sender Is Button1 Then
                Clean_Data()
                Select Case rqProecess
                    Case "addmyall"
                        Call Add_Student()
                        Table4.Style.Item("display") = "none"
                    Case "addother"
                        Call Clean_Data()
                        Dim rstA2 As Boolean = Actoin_cnt2(rqDLID)
                        If rstA2 Then
                            Page.RegisterStartupScript(TIMS.xBlockName, "<script>location.href='SD_05_008.aspx?ID=" & s_MRqID & "';</script>")
                        End If

                        'Dim iResultCount As Integer = 0
                        'Dim iStdCount As Integer = 0
                        'sql = "SELECT ResultCount FROM STUD_DATALID WHERE DLID='" & rqDLID & "'"
                        'ResultCount = DbAccess.ExecuteScalar(sql, objconn)
                        'sql = "SELECT COUNT(1) FROM STUD_RESULTSTUDDATA WHERE DLID='" & rqDLID & "'"
                        'StdCount = DbAccess.ExecuteScalar(sql, objconn)
                        'If StdCount >= ResultCount Then
                        '    Common.MessageBox(Me, "已達到結訓人數，不能在新增!")
                        '    If Session(cst_search) Is Nothing Then Session(cst_search) = Me.ViewState(cst_search)

                        '    Page.RegisterStartupScript("rrr", "<script>location.href='SD_05_008.aspx?ID=" & Request("ID") & "';</script>")
                        'End If
                    Case "add", "edit"
                        If Session(cst_search) Is Nothing Then Session(cst_search) = Me.ViewState(cst_search)

                        Common.RespWrite(Me, "<script>")
                        Common.RespWrite(Me, "alert('" & sMsgbox & "');")
                        Common.RespWrite(Me, "window.location='SD_05_008.aspx?ID=" & s_MRqID & "'")
                        Common.RespWrite(Me, "</script>")
                End Select
            Else
                If Session(cst_search) Is Nothing Then Session(cst_search) = Me.ViewState(cst_search)

                Common.RespWrite(Me, "<script>")
                Common.RespWrite(Me, "alert('" & sMsgbox & "');")
                Common.RespWrite(Me, "window.location='SD_05_008.aspx?ID=" & s_MRqID & "'")
                Common.RespWrite(Me, "</script>")
            End If
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Call TIMS.CloseDbConn(tConn)

            sMsgbox += "失敗"
            'Common.RespWrite(Me, ex)
            Throw ex
        End Try
        Call TIMS.CloseDbConn(tConn)

        '結束
        If sMsgbox <> "" Then Common.MessageBox(Me, sMsgbox)

    End Sub

    '回查詢頁面
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Session(cst_search) Is Nothing Then Session(cst_search) = Me.ViewState(cst_search)

        TIMS.Utl_Redirect1(Me, "SD_05_008.aspx?ID=" & TIMS.Get_MRqID(Me))
    End Sub

    Public Shared Function Check_QdataDr(ByVal OCID As String, ByVal SOCID As String, ByRef conn As SqlConnection) As DataRow
        Dim rst As Boolean = False 'False:有正常填寫。true:有遺漏
        SOCID = TIMS.ClearSQM(SOCID)
        Dim sql As String = ""
        sql &= " select b.DLID" & vbCrLf
        sql &= " ,b.Q7" & vbCrLf
        sql &= " ,b.Q8,b.Q8OTHER" & vbCrLf
        sql &= " ,b.Q9,b.Q9Y" & vbCrLf
        sql &= " ,b.Q10,b.Q11,b.Q11N" & vbCrLf
        sql &= " ,b.Q12A,b.Q12B" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN STUD_RESULTSTUDDATA b on a.socid = b.socid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.StudStatus not in (2,3) " & vbCrLf '非離退訓。
        sql &= " AND a.OCID='" & OCID & "'" & vbCrLf
        sql &= " and a.SOCID ='" & SOCID & "'" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sql, conn)
        Return dr
    End Function

    ''' <summary>'檢查是有有填問卷及是否有題目 未填完整 紅色【*】表示:此班級有學員還沒填寫問卷或是填寫的問卷作答有遺漏</summary>
    ''' <param name="ocid"></param>
    ''' <param name="SOCID"></param>
    ''' <param name="conn"></param>
    ''' <returns> 'False:'表示都沒有未填寫的 ／True:'表示有未填寫的</returns>
    Public Shared Function Check_Qdata(ByVal OCID As String, ByVal SOCID As String, ByRef conn As SqlConnection) As Boolean
        Dim rst As Boolean = False 'False:有正常填寫。true:有遺漏
        SOCID = TIMS.ClearSQM(SOCID)

        Dim PMS1 As New Hashtable From {{"OCID", TIMS.CINT1(OCID)}}
        Dim sql As String = ""
        sql &= " select b.DLID" & vbCrLf
        sql &= " ,b.Q7" & vbCrLf
        sql &= " ,b.Q8,b.Q8OTHER" & vbCrLf
        sql &= " ,b.Q9,b.Q9Y" & vbCrLf
        sql &= " ,b.Q10,b.Q11,b.Q11N" & vbCrLf
        sql &= " ,b.Q12A,b.Q12B" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " LEFT JOIN STUD_RESULTSTUDDATA b on a.socid = b.socid" & vbCrLf
        sql &= " WHERE a.StudStatus not in (2,3) " & vbCrLf '非離退訓。
        sql &= " AND a.OCID=@OCID" & vbCrLf
        If SOCID <> "" Then
            PMS1.Add("SOCID", TIMS.CINT1(SOCID))
            sql &= " and a.SOCID=@SOCID" & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, conn, PMS1)

        For Each dr As DataRow In dt.Rows
            rst = False
            If Convert.ToString(dr("Q7")) = "" Then rst = True
            If Convert.ToString(dr("Q8")) = "" Then rst = True
            If Convert.ToString(dr("Q9")) = "" Then rst = True
            If Convert.ToString(dr("Q10")) = "" Then rst = True
            If Convert.ToString(dr("Q8")) = "5" AndAlso Convert.ToString(dr("Q8Other")) = "" Then rst = True
            If Convert.ToString(dr("Q9")) = "Y" AndAlso Convert.ToString(dr("Q9Y")) = "" Then rst = True
            'If Convert.ToString(dr("Q8")) <> "4" AndAlso Convert.ToString(dr("Q11N")) = "" Then rst = True
            If Convert.ToString(dr("Q11")) = "N" AndAlso Convert.ToString(dr("Q11N")) = "" Then rst = True

            If rst Then Exit For
        Next
        'False:'表示都沒有未填寫的 ／True:'表示有未填寫的
        Return rst
    End Function

    Function Actoin_cnt2(ByVal rqDLID As String) As Boolean
        Dim rst As Boolean = False

        Dim sql As String = ""
        Dim iResultCount As Integer = 0
        Dim iStdCount As Integer = 0
        sql = "SELECT ResultCount FROM STUD_DATALID WHERE DLID='" & rqDLID & "'"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then
            iResultCount = Val(dr("ResultCount"))
        End If
        'iResultCount = DbAccess.ExecuteScalar(sql, objconn)
        sql = "SELECT COUNT(1) cnt FROM STUD_RESULTSTUDDATA WHERE DLID='" & rqDLID & "'"
        iStdCount = DbAccess.ExecuteScalar(sql, objconn)
        If iStdCount >= iResultCount Then
            Common.MessageBox(Me, "已達到結訓人數，不能再繼續新增!")

            If Session(cst_search) Is Nothing Then Session(cst_search) = Me.ViewState(cst_search)
            rst = True
            'Page.RegisterStartupScript(TIMS.xBlockName, "<script>location.href='SD_05_008.aspx?ID=" & Request("ID") & "';</script>")
        End If
        Return rst
    End Function

End Class
