Public Class SD_05_031
    Inherits AuthBasePage ' Global.System.Web.UI.Page

    Const CST_ERRMSG3 As String = "就服資料庫連線異常!"
    Dim ff3 As String = ""
    Dim objconn As SqlConnection
    Dim objconn3 As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As Global.System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
        Call TIMS.CloseDbConn(objconn3)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As Global.System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'Dim strConn3 As String = ""
        Dim strConn3 As String = TIMS.Utl_GetConfigSet("ConnectionEvtadb1Enc")
        If strConn3 <> "" Then strConn3 = RSA20031.AesDecrypt2(strConn3)
        If strConn3 = "" Then strConn3 = TIMS.Utl_GetConfigSet("ConnectionEvtadb1")
        'strConn3 = ConfigurationSettings.AppSettings("ConnectionEvtadb1")
        objconn3 = New SqlConnection(strConn3)

        Try
            Call TIMS.OpenDbConn(objconn3)
        Catch ex As Exception
            objconn3 = Nothing
            'Common.MessageBox(Me, "就服資料庫連線異常!!")
            'Exit Sub
        End Try

        If Not IsPostBack Then
            'Me.sIDNO.Text = "K201508847"
        End If

    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""
        If sIDNO.Text = "" AndAlso sName.Text = "" AndAlso sbirthday.Text = "" Then
            Errmsg &= "請輸入任一條件" & vbCrLf
        End If

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    '查詢
    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call GetStudentData()
    End Sub

    Sub GetStudentData()
        'Dim SearchStr1 As String
        'Dim SearchStr2 As String
        'Dim SearchStr3 As String

        Dim dt1 As DataTable
        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " SELECT TOP 10" & vbCrLf
        Sql += "        ss.name" & vbCrLf
        Sql += "        ,ss.idno" & vbCrLf
        Sql += "        ,CONVERT(varchar, ss.birthday, 111) birthday" & vbCrLf
        Sql += "        ,s2.phoned" & vbCrLf
        Sql += "        ,iz.ZipName" & vbCrLf
        Sql += "        ,s2.Address Address" & vbCrLf
        Sql += " FROM stud_studentinfo ss" & vbCrLf
        Sql += " JOIN stud_subdata s2 ON s2.sid = ss.sid" & vbCrLf
        Sql += " JOIN view_zipname iz ON iz.zipCode = s2.zipCode1" & vbCrLf
        Sql += " WHERE 1=1" & vbCrLf

        If sIDNO.Text <> "" Then Sql += " AND ss.idno LIKE '" & sIDNO.Text & "%'" & vbCrLf
        If sName.Text <> "" Then Sql += " AND ss.name LIKE '" & sName.Text & "%'" & vbCrLf
        If sbirthday.Text <> "" Then Sql += " AND ss.birthday = " & TIMS.To_date(sbirthday.Text) & vbCrLf
        dt1 = DbAccess.GetDataTable(Sql, objconn)

        msg.Text = "查無資料"
        Me.Panelshow1.Visible = False
        If dt1.Rows.Count > 0 Then
            Dim dr As DataRow = dt1.Rows(0)
            msg.Text = ""
            Me.Panelshow1.Visible = True
            labidno.Text = dr("idno")
            labname.Text = dr("name")
            'labidno.Text =dr("idno")
            labbirthday.Text = dr("birthday")
            labtel.Text = dr("phoned")
            labaddress.Text = dr("zipname") & dr("address")
        End If

        If objconn3 Is Nothing Then
            msg.Text = "就服資料庫連線異常"
            Me.Panelshow1.Visible = False
        End If

        If dt1.Rows.Count = 0 AndAlso Not objconn3 Is Nothing Then
            Dim odt As New DataTable
            '(連接 "就服" Oracle DB, 所以不用調整SQL語法)
            Sql = "" & vbCrLf
            Sql += " SELECT a.idno" & vbCrLf
            Sql += "        ,a.name" & vbCrLf
            Sql += "        ,TO_CHAR(a.birth, 'YYYY/MM/DD') AS birthday" & vbCrLf
            Sql += "        ,a.tel1 phoned" & vbCrLf
            Sql += "        ,a.addr address" & vbCrLf
            Sql += " FROM apltbl a" & vbCrLf
            Sql += " where rownum <=10" & vbCrLf '不用太多資料只要1筆

            If sIDNO.Text <> "" Then Sql += " AND a.idno = @IDNO" & vbCrLf
            If sName.Text <> "" Then Sql += " AND a.name LIKE '" & sName.Text & "%'" & vbCrLf
            If sbirthday.Text <> "" Then Sql += " AND a.birth = " & TIMS.To_date(sbirthday.Text) & vbCrLf
            Dim cmd As New SqlCommand(Sql, objconn3)

            Try
                With cmd
                    .Parameters.Clear()
                    If sIDNO.Text <> "" Then
                        .Parameters.Add("IDNO", SqlDbType.VarChar).Value = sIDNO.Text
                    End If
                    odt.Load(.ExecuteReader())
                End With
            Catch ex As Exception
                objconn3 = Nothing
                'Common.MessageBox(Me, "就服資料庫連線異常!!")
                'Exit Sub
            End Try

            If objconn3 Is Nothing Then Common.MessageBox(Me, "就服資料庫連線異常!!")
            msg.Text = "查無資料"
            Me.Panelshow1.Visible = False
            If odt.Rows.Count > 0 Then
                Dim dr As DataRow = odt.Rows(0)
                msg.Text = ""
                Me.Panelshow1.Visible = True
                labidno.Text = Convert.ToString(dr("idno"))
                labname.Text = Convert.ToString(dr("name"))
                'labidno.Text =dr("idno")
                labbirthday.Text = Convert.ToString(dr("birthday"))
                labtel.Text = Convert.ToString(dr("phoned"))
                labaddress.Text = Convert.ToString(dr("address"))
            End If
        End If

        [labmsg1].Text = "查無資料!"
        [labmsg2].Text = "查無資料!"
        [labmsg3].Text = "查無資料!"
        [Labmsg4].Text = "查無資料!"
        [Labmsg5].Text = "查無資料!"
        [Labmsg6].Text = "查無資料!"

        If labidno.Text <> "" Then
            Call showList1(labidno.Text)
            Call showList2(labidno.Text)
            Call showList3(labidno.Text)
            Call showList4(labidno.Text)
            Call showList5(labidno.Text)
            Call showList6(labidno.Text)
        End If
    End Sub

    Sub showList1(ByVal vIDNO As String)
        Dim SearchStr1 As String = " AND a.SID ='" & vIDNO & "'"
        Dim SearchStr2 As String = " AND a.IDNO='" & vIDNO & "'"
        Dim SearchStr3 As String = " AND b.IDNO='" & vIDNO & "'"

        Dim RecordCountInt As Integer = 2000 '最大筆數限制
        Dim Key_Identity As DataTable
        Key_Identity = TIMS.Get_KeyTable("Key_Identity", "", objconn)

        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim dr1 As DataRow
        Dim dr2 As DataRow
        'Dim dr3 As DataRow

        '建立DataGird用的DataTable格式 Start
        dt = New DataTable
        dt.Columns.Add(New DataColumn("IDNO"))
        dt.Columns.Add(New DataColumn("Name"))
        dt.Columns.Add(New DataColumn("Sex"))
        dt.Columns.Add(New DataColumn("Birthday", System.Type.GetType("System.DateTime")))
        dt.Columns.Add(New DataColumn("DistName"))                  '轄區分署(轄區中心)
        dt.Columns.Add(New DataColumn("Years"))
        dt.Columns.Add(New DataColumn("PlanName"))
        dt.Columns.Add(New DataColumn("OrgName"))                   '訓練機構
        dt.Columns.Add(New DataColumn("TMID"))                      '訓練職類
        dt.Columns.Add(New DataColumn("CJOB_NAME"))                 '通俗職類
        dt.Columns.Add(New DataColumn("ClassName"))                 '班別
        dt.Columns.Add(New DataColumn("THours"))                    '受訓時數
        dt.Columns.Add(New DataColumn("TRound"))                    '受訓期間
        dt.Columns.Add(New DataColumn("SkillName"))                 '技能檢定
        dt.Columns.Add(New DataColumn("TFlag"))                     '訓練狀態
        dt.Columns.Add(New DataColumn("Ident"))                     '身分別
        dt.Columns.Add(New DataColumn("Tel"))                       '聯絡電話
        dt.Columns.Add(New DataColumn("Address"))                   '聯絡地址
        dt.Columns.Add(New DataColumn("WEEKS"))                     '上課地點
        '建立DataGird用的DataTable格式 End

        sql = ""
        sql += " SELECT a.SID " 'IDNO
        sql += "        ,a.Name "
        sql += "        ,a.Sex "
        sql += "        ,a.Birth "
        sql += "        ,a.DistName "
        sql += "        ,a.Years "
        sql += "        ,a.PlanName "
        sql += "        ,a.TrinUnit "
        sql += "        ,a.ClassName "
        sql += "        ,a.SDate "
        sql += "        ,a.EDate "
        sql += "        ,a.Ident "
        sql += "        ,a.Tel "
        sql += "        ,a.Addr "
        sql += " FROM StdAll a WHERE 1=1 " & SearchStr1

        Try
            'dt1 = DbAccess.GetDataTable(sql, objconn)
            Dim cmd As New SqlCommand(sql, objconn)
            With cmd
                .Parameters.Clear()
                '.Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
                dt1.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try

        For Each dr1 In dt1.Rows
            If RecordCountInt > 0 Then
                RecordCountInt -= 1
            Else
                Exit For '超過 最大筆數限制
            End If

            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("IDNO") = TIMS.ChangeIDNO(dr1("SID"))
            dr("Name") = dr1("Name")
            dr("Sex") = dr1("Sex")
            dr("Birthday") = dr1("Birth")
            dr("DistName") = dr1("DistName")
            dr("Years") = dr1("Years")
            dr("PlanName") = dr1("PlanName")
            dr("OrgName") = dr1("TrinUnit")
            'dr("TMID") = dr1("")
            dr("ClassName") = dr1("ClassName")
            'dr("THours") = dr1("")
            If dr1("SDate").ToString <> "" And dr1("EDate").ToString <> "" Then
                dr("TRound") = FormatDateTime(dr1("SDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr1("EDate"), DateFormat.ShortDate)
            End If
            'dr("SkillName") = dr1("")
            dr("TFlag") = "結訓"
            dr("Ident") = IIf(IsNumeric(dr1("Ident")), "無法辨別", dr1("Ident").ToString)
            dr("Tel") = dr1("Tel").ToString
            dr("Address") = dr1("Addr").ToString
        Next

        sql = ""
        sql += " SELECT a.Serial "
        sql += "        ,a.IDNO "
        sql += "        ,a.Name "
        sql += "        ,a.Sex "
        sql += "        ,a.Birth "
        sql += "        ,a.DistName "
        sql += "        ,a.PlanName "
        sql += "        ,a.TrinUnit "
        sql += "        ,a.ClassName "
        sql += "        ,a.SDate "
        sql += "        ,a.EDate "
        sql += "        ,a.Ident "
        sql += "        ,a.Tel "
        sql += "        ,a.Addr "
        sql += "        ,b.TrainName "
        sql += " FROM History_StudentInfo93 a "
        sql += " LEFT JOIN Key_TrainType b ON a.TMID = b.TMID "
        sql += " WHERE 1=1" & vbCrLf
        sql += SearchStr2

        Try
            'dt2 = DbAccess.GetDataTable(sql, objconn)
            Dim cmd As New SqlCommand(sql, objconn)
            With cmd
                .Parameters.Clear()
                '.Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
                dt2.Load(.ExecuteReader())
            End With
            'dt2 = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try

        For Each dr2 In dt2.Rows
            If RecordCountInt > 0 Then
                RecordCountInt -= 1
            Else
                Exit For '超過 最大筆數限制
            End If
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("IDNO") = TIMS.ChangeIDNO(dr2("IDNO"))
            dr("Name") = dr2("Name")
            dr("Sex") = dr2("Sex")
            dr("Birthday") = dr2("Birth")
            dr("DistName") = dr2("DistName")
            dr("PlanName") = dr2("PlanName")
            dr("OrgName") = dr2("TrinUnit")
            dr("TMID") = dr2("TrainName") 'b.TrainName
            dr("ClassName") = dr2("ClassName")
            dr("TRound") = FormatDateTime(dr2("SDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr2("EDate"), DateFormat.ShortDate)
            dr("TFlag") = "結訓"
            If Key_Identity.Select("IdentityID='" & dr2("Ident") & "'").Length > 0 Then
                dr("Ident") = Key_Identity.Select("IdentityID='" & dr2("Ident") & "'")(0)("Name")
            Else
                dr("Ident") = "無身分別"
            End If
            dr("Tel") = dr2("Tel").ToString
            dr("Address") = dr2("Addr").ToString
        Next

        sql = ""
        sql += " SELECT b.IDNO ,b.Name ,b.Sex ,b.Birthday" & vbCrLf
        sql += "  ,f.Name AS DistName ,e.OrgName ,g.TrainName TMID" & vbCrLf
        'sql += "        ,c.ClassCName +'第' +c.cyclType +'期' AS ClassName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(c.CLASSCNAME,c.CYCLTYPE) CLASSNAME" & vbCrLf
        sql += " ,CASE WHEN a.TrainHours IS NULL THEN c.THours ELSE a.TrainHours END THours" & vbCrLf
        'sql += "       ,CASE WHEN a.OpenDate IS NULL THEN c.STDate ELSE a.OpenDate END STDate" & vbCrLf      '2009/08/25 改成以班級的開結訓日為開結訓日
        'sql += "       ,CASE WHEN a.CloseDate IS NULL THEN c.FTDate ELSE a.CloseDate END FTDate" & vbCrLf
        sql += " ,c.STDate" & vbCrLf
        sql += " ,c.FTDate" & vbCrLf
        sql += " ,a.TrainHours ,a.RejectTDate1, a.RejectTDate2" & vbCrLf
        sql += " ,h.ExamName ,a.StudStatus ,a.MIdentityID" & vbCrLf
        sql += " ,j.PhoneD ,j.ZipCode1 ,j.Address ,k.PlanName ,i.Years ,s.CJOB_NAME" & vbCrLf
        sql += " ,dbo.fn_GET_PLAN_ONCLASS2(pp.PlanID,pp.ComIDNO,pp.SeqNo,'WEEKTIME') WEEKS" & vbCrLf
        sql += " FROM Class_StudentsOfClass a" & vbCrLf
        sql += " JOIN Stud_StudentInfo b ON a.SID = b.SID" & vbCrLf
        sql += " JOIN Class_ClassInfo c ON a.OCID = c.OCID" & vbCrLf
        sql += " JOIN Plan_PlanInfo pp ON c.planid = pp.planid AND pp.comidno = c.comidno AND pp.seqno = c.seqno" & vbCrLf
        sql += " JOIN Auth_Relship d ON c.RID = d.RID" & vbCrLf
        sql += " JOIN Org_OrgInfo e ON d.OrgID = e.OrgID" & vbCrLf
        sql += " LEFT JOIN ID_District f ON d.DistID = f.DistID" & vbCrLf
        sql += " LEFT JOIN Key_TrainType g ON c.TMID = g.TMID" & vbCrLf
        sql += " LEFT JOIN SHARE_CJOB s ON s.CJOB_UNKEY = c.CJOB_UNKEY" & vbCrLf
        sql += " LEFT JOIN Stud_TechExam h ON a.SOCID = h.SOCID" & vbCrLf
        sql += " JOIN ID_Plan i  ON i.PlanID = c.PlanID" & vbCrLf
        sql += " JOIN KEY_plan k  ON k.TPlanID = i.TPlanID" & vbCrLf
        sql += " JOIN Stud_SubData j  ON j.SID = b.SID" & vbCrLf
        sql += " Where 1=1" & SearchStr3

        Try
            'dt3 = DbAccess.GetDataTable(sql, objconn)
            Dim cmd As New SqlCommand(sql, objconn)
            With cmd
                .Parameters.Clear()
                '.Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
                dt3.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try

        For Each dr3 As DataRow In dt3.Rows
            If RecordCountInt > 0 Then
                RecordCountInt -= 1
            Else
                Exit For '超過 最大筆數限制
            End If
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("IDNO") = TIMS.ChangeIDNO(dr3("IDNO"))
            dr("Name") = dr3("Name")
            dr("Sex") = dr3("Sex")
            dr("Birthday") = dr3("Birthday")
            dr("DistName") = dr3("DistName")
            dr("Years") = dr3("Years")
            dr("PlanName") = dr3("PlanName")
            dr("OrgName") = dr3("OrgName")
            dr("TMID") = dr3("TMID")
            dr("CJOB_NAME") = dr3("CJOB_NAME")
            dr("ClassName") = dr3("ClassName")
            dr("WEEKS") = dr3("WEEKS")

            Select Case dr3("StudStatus").ToString '訓練狀態，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                Case "2" '"離訓"
                    dr("THours") = "<FONT color='Red'>" & dr3("TrainHours") & "</FONT>" '參訓時數，以 Class_StudentsOfClass 為主
                    Try
                        dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr3("RejectTDate1"), DateFormat.ShortDate)
                    Catch ex As Exception
                        dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>離訓日期異常"
                    End Try
                Case "3" '"退訓"
                    dr("THours") = "<FONT color='Red'>" & dr3("TrainHours") & "</FONT>" '參訓時數，以 Class_StudentsOfClass 為主
                    Try
                        dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr3("RejectTDate2"), DateFormat.ShortDate)
                    Catch ex As Exception
                        dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>退訓日期異常"
                    End Try
                Case Else
                    dr("THours") = dr3("THours") '參訓時數，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                    dr("TRound") = FormatDateTime(dr3("STDate"), DateFormat.ShortDate) & "<BR>|<BR>" & FormatDateTime(dr3("FTDate"), DateFormat.ShortDate)
            End Select

            dr("SkillName") = dr3("ExamName")
            Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(dr3("StudStatus"))
            dr("TFlag") = STUDSTATUS_N '"在訓"

            ff3 = "IdentityID='" & dr3("MIdentityID") & "'"
            If Key_Identity.Select(ff3).Length > 0 Then
                dr("Ident") = Key_Identity.Select(ff3)(0)("Name")
            Else
                dr("Ident") = "無身分別"
            End If
            dr("Tel") = dr3("PhoneD").ToString
            dr("Address") = TIMS.Get_ZipName(dr3("ZipCode1")) & dr3("Address").ToString
        Next
        [labmsg1].Text = "查無資料!"
        If dt.Rows.Count > 0 Then
            [labmsg1].Text = ""
            dt.DefaultView.Sort = "IDNO,Birthday,TRound"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    Sub showList2(ByVal IDNO As String)
        'param.Add("IDNO", IDNO)
        'dt = db.QueryForDataTableAll("SelSD01004S16", param)

        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.IDNO,e.OrgName" & vbCrLf
        'sql += " ,d.ClassCName + '第' + d.CyclType + '期' ClassCName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(d.CLASSCNAME,d.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql += " ,d.STDate,d.FTDate" & vbCrLf
        sql += " ,a.TrainingMoney" & vbCrLf
        sql += " ,a.PayMoney" & vbCrLf
        sql += " FROM sub_subsidyapply a" & vbCrLf
        sql += " JOIN Class_StudentsOfClass b ON a.SOCID = b.SOCID" & vbCrLf
        sql += " JOIN Stud_StudentInfo c ON b.SID = c.SID" & vbCrLf
        sql += " JOIN Class_ClassInfo d ON b.OCID = d.OCID" & vbCrLf
        sql += " JOIN view_RIDName e ON d.RID = e.RID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += "    AND c.IDNO = @IDNO" & vbCrLf

        Dim dt As New DataTable
        Dim cmd As New SqlCommand(sql, objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
            dt.Load(.ExecuteReader())
        End With

        [labmsg2].Text = "查無資料!"
        If dt.Rows.Count > 0 Then
            [labmsg2].Text = ""
            DataGrid2.DataSource = dt
            DataGrid2.DataBind()
        End If
    End Sub

    '僱用與否
    Function funShareName(ByVal share_type As String) As String
        If objconn3 Is Nothing Then Return ""
        Dim rst As String = ""
        Dim odt As New DataTable
        Dim sql As String = ""
        sql = " SELECT share_name FROM share_source WHERE share_type = @share_type AND share_id = '10' "
        Dim cmd As New SqlCommand(sql, objconn3)

        With cmd
            .Parameters.Clear()
            .Parameters.Add("share_type", SqlDbType.VarChar).Value = share_type
            odt.Load(.ExecuteReader())
        End With

        If odt.Rows.Count > 0 Then rst = odt.Rows(0)("share_name")
        Return rst
    End Function

    '職業名稱
    Function funOccucd(ByVal job_no As String) As String
        If objconn3 Is Nothing Then Return ""
        Dim rst As String = ""
        Dim odt As New DataTable
        Dim sql As String = ""
        sql = " SELECT job_name FROM share_job WHERE job_no = @job_no "
        Dim cmd As New SqlCommand(sql, objconn3)

        With cmd
            .Parameters.Clear()
            .Parameters.Add("job_no", SqlDbType.VarChar).Value = job_no
            odt.Load(.ExecuteReader())
        End With

        If odt.Rows.Count > 0 Then rst = odt.Rows(0)("job_name")
        Return rst
    End Function

    '以介紹檔(INTRTLL)的建置日期是否在2005/12/31之後，判定使用新版或舊版的回覆畫面
    Function IsNewIntrtbll(ByVal unkey As String) As Boolean
        If objconn3 Is Nothing Then Return ""
        Dim rst As Boolean = False '舊
        Dim odt As New DataTable
        Dim sql As String = ""
        'sql = " SELECT CREATE_DATE FROM INTRTBLL WHERE INTR_UNKEY = @unkey AND to_date(CREATE_DATE,'yyyy-MM-dd hh24.mi.ss')-to_date('20051231','yyyymmdd')>0 "
        sql = " SELECT CREATE_DATE FROM INTRTBLL WHERE INTR_UNKEY = @unkey AND DATEDIFF(DAY, CONVERT(DATETIME, '2005/12/31'), CREATE_DATE) >= 0 "  'edit，by:20181101
        Dim cmd As New SqlCommand(sql, objconn3)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("unkey", SqlDbType.VarChar).Value = unkey
            odt.Load(.ExecuteReader())
        End With
        If odt.Rows.Count > 0 Then
            rst = True
        End If
        Return rst
    End Function

    '雇主未能僱用原因
    Function funReplyNonFill(ByVal unkey As String, ByVal reply_ansres As String, ByVal reply_non_intro As String, ByVal reply_non_intro_other As String) As String
        If objconn3 Is Nothing Then Return ""
        Dim strRet As String = ""
        Dim sql As String = ""
        Dim odt As New DataTable
        Dim cmd As SqlCommand = Nothing
        Dim isNew As Boolean = IsNewIntrtbll(unkey)

        If (reply_ansres <> "9") Then
            If isNew Then
                If (reply_non_intro <> "A") Then
                    '依政府新規定年齡不合及性別不合不顯示
                    If (reply_non_intro <> "6" AndAlso reply_non_intro <> "7") Then
                        sql = " SELECT SHARE_NAME FROM SHARE_SOURCE WHERE SHARE_TYPE= '212' AND SHARE_ID = @reply_non_intro "
                        cmd = New SqlCommand(sql, objconn3)
                        With cmd
                            .Parameters.Clear()
                            .Parameters.Add("reply_non_intro", SqlDbType.VarChar).Value = reply_non_intro
                            odt = New DataTable
                            odt.Load(.ExecuteReader())
                        End With
                        If odt.Rows.Count > 0 Then strRet = odt.Rows(0)("SHARE_NAME")
                    End If
                Else
                    strRet = reply_non_intro_other
                End If
            Else
                If (reply_non_intro <> "A") Then
                    '依政府新規定年齡不合及性別不合不顯示
                    If (reply_non_intro <> "8" AndAlso reply_non_intro <> "9") Then
                        sql = " SELECT SHARE_NAME FROM SHARE_SOURCE WHERE SHARE_TYPE = '12' AND SHARE_ID = @reply_non_intro "
                        cmd = New SqlCommand(sql, objconn3)
                        With cmd
                            .Parameters.Clear()
                            .Parameters.Add("reply_non_intro", SqlDbType.VarChar).Value = reply_non_intro
                            odt = New DataTable
                            odt.Load(.ExecuteReader())
                        End With
                        If odt.Rows.Count > 0 Then strRet = odt.Rows(0)("SHARE_NAME")
                    End If
                Else
                    strRet = reply_non_intro_other
                End If
            End If
        End If
        Return strRet
    End Function

    '求職未能推介原因
    Function funReplyNonIntro(ByVal unkey As String, ByVal reply_ansres As String, ByVal reply_non_intro As String, ByVal reply_non_intro_other As String) As String
        If objconn3 Is Nothing Then Return ""

        Dim strRet As String = ""
        Dim sql As String = ""
        Dim odt As New DataTable
        Dim cmd As SqlCommand = Nothing
        Dim isNew As Boolean = IsNewIntrtbll(unkey)

        If (reply_ansres <> "9") Then
            If isNew Then
                If (reply_non_intro <> "A") Then
                    '依政府新規定年齡不合及性別不合不顯示
                    If (reply_non_intro <> "7" AndAlso reply_non_intro <> "8") Then
                        sql = " SELECT SHARE_NAME FROM SHARE_SOURCE WHERE SHARE_TYPE = '211' AND SHARE_ID = @reply_non_intro "
                        cmd = New SqlCommand(sql, objconn3)
                        With cmd
                            .Parameters.Clear()
                            .Parameters.Add("reply_non_intro", SqlDbType.VarChar).Value = reply_non_intro
                            odt = New DataTable
                            odt.Load(.ExecuteReader())
                        End With
                        If odt.Rows.Count > 0 Then strRet = odt.Rows(0)("SHARE_NAME")
                    End If
                Else
                    strRet = reply_non_intro_other
                End If
            Else
                If (reply_non_intro <> "A") Then
                    '依政府新規定年齡不合及性別不合不顯示
                    If (reply_non_intro <> "7" AndAlso reply_non_intro <> "8") Then
                        sql = " SELECT SHARE_NAME FROM SHARE_SOURCE WHERE SHARE_TYPE = '11' AND SHARE_ID = @reply_non_intro "
                        cmd = New SqlCommand(sql, objconn3)
                        With cmd
                            .Parameters.Clear()
                            .Parameters.Add("reply_non_intro", SqlDbType.VarChar).Value = reply_non_intro
                            odt = New DataTable
                            odt.Load(.ExecuteReader())
                        End With
                        If odt.Rows.Count > 0 Then
                            strRet = odt.Rows(0)("SHARE_NAME")
                        End If
                    End If
                Else
                    strRet = reply_non_intro_other
                End If
            End If
        End If
        Return strRet
    End Function

    '就業服務(推介歷程) 就服oracle
    Sub showList3(ByVal vIDNO As String)
        If objconn3 Is Nothing Then [labmsg3].Text = CST_ERRMSG3 '"就服資料庫連線異常!"
        If objconn3 Is Nothing Then Exit Sub

        Try
            DbAccess.Open(objconn3)
        Catch ex As Exception
            [labmsg3].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try

        Dim odt As New DataTable
        '(連接 "就服" Oracle DB, 所以不用調整SQL語法)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.idno" & vbCrLf
        sql += "  ,a.name" & vbCrLf
        sql += "  ,a.birth" & vbCrLf
        sql += "  ,j.intr_unkey" & vbCrLf
        sql += "  ,TO_CHAR(j.intr_irgdate, 'YYYY/MM/DD') AS interDate1" & vbCrLf
        'sql += " --,TO_CHAR(add_months(getdate(),-24), 'YYYY/MM/DD') AS t2" & vbCrLf
        sql += "  ,c.compname OrgName" & vbCrLf
        'sql += " --,j.intr_employer_id" & vbCrLf
        'sql += " --,j.intr_hire_id" & vbCrLf
        sql += "  ,j.intr_occucd Occucd" & vbCrLf
        sql += "  ,j.reply_ansres" & vbCrLf
        'sql += " ,j.reply_non_fill" & vbCrLf
        'sql += " ,j.reply_non_fill_other" & vbCrLf
        sql += "  ,j.reply_non_intro" & vbCrLf
        sql += "  ,j.reply_non_intro_other" & vbCrLf
        sql += " FROM intrtbll j" & vbCrLf
        sql += " JOIN apltbl a ON j.intr_labor_id = a.labor_id" & vbCrLf
        sql += " JOIN vnbtbl c ON c.idno = j.intr_idno2" & vbCrLf
        sql += " WHERE j.del_state = 0" & vbCrLf
        sql += " AND a.IDNO = @IDNO" & vbCrLf
        'sql += "   --AND j.intr_irgdate >= TO_CHAR(add_months(getdate(),-24), 'YYYY-MM-DD')" & vbCrLf
        sql += " AND j.intr_irgdate >= trunc(add_months(getdate(),-24))" & vbCrLf
        'sql += "   AND rownum <= 10" & vbCrLf
        sql += " ORDER BY j.intr_irgdate DESC" & vbCrLf
        Dim cmd As New SqlCommand(sql, objconn3)

        Try
            With cmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = vIDNO
                odt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            [labmsg3].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try

        '建立DataGird用的DataTable格式 Start
        Dim dt As DataTable
        dt = New DataTable
        dt.Columns.Add(New DataColumn("seqno"))
        dt.Columns.Add(New DataColumn("interDate1"))
        dt.Columns.Add(New DataColumn("OrgName"))
        dt.Columns.Add(New DataColumn("workName"))
        dt.Columns.Add(New DataColumn("workYN"))
        dt.Columns.Add(New DataColumn("NRReason"))
        dt.Columns.Add(New DataColumn("NWReason"))
        '建立DataGird用的DataTable格式 End

        For Each odr As DataRow In odt.Rows
            Dim dr As DataRow = dt.NewRow
            dt.Rows.Add(dr)
            dr("seqno") = dt.Rows.Count
            dr("interDate1") = Convert.ToString(odr("interDate1")) ' "2012/07/18"
            dr("OrgName") = Convert.ToString(odr("OrgName")) '"基隆港務局"
            dr("workName") = funOccucd(Convert.ToString(odr("Occucd"))) '"其他環境清潔維護工及有關體力工"
            dr("workYN") = funShareName(Convert.ToString(odr("reply_ansres"))) '"介紹錄用"
            dr("NRReason") = funReplyNonFill(Convert.ToString(odr("intr_unkey")), Convert.ToString(odr("reply_ansres")), Convert.ToString(odr("reply_non_intro")), Convert.ToString(odr("reply_non_intro_other"))) '""'labReplyAnsres
            dr("NWReason") = funReplyNonIntro(Convert.ToString(odr("intr_unkey")), Convert.ToString(odr("reply_ansres")), Convert.ToString(odr("reply_non_intro")), Convert.ToString(odr("reply_non_intro_other"))) '""'labReplyAnsres2
        Next
        'E121317434

#Region "(No Use)"

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("interDate1") = "2012/07/18"
        'dr("OrgName") = "基隆港務局"
        'dr("workName") = "其他環境清潔維護工及有關體力工"
        'dr("workYN") = "介紹錄用"
        'dr("NRReason") = ""
        'dr("NWReason") = ""

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("interDate1") = "2012/01/09"
        'dr("OrgName") = "何朝榮診所"
        'dr("workName") = "工讀生"
        'dr("workYN") = "介紹未錄用"
        'dr("NRReason") = "求才廠商已錄用滿額"
        'dr("NWReason") = ""

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("interDate1") = "2012/01/02"
        'dr("OrgName") = "阿瘦實業-基隆店"
        'dr("workName") = "零售業務員"
        'dr("workYN") = "介紹未錄用"
        'dr("NRReason") = "未備駕照或交通工具"
        'dr("NWReason") = ""

#End Region

        'dt.AcceptChanges()
        [labmsg3].Text = "查無資料!"
        If dt.Rows.Count > 0 Then
            [labmsg3].Text = ""
            'dt.DefaultView.Sort = "IDNO,Birthday,TRound"
            DataGrid3.DataSource = dt
            DataGrid3.DataBind()
        End If

        'DbAccess.Close(objconn3)
    End Sub

    '技能檢定。就服oracle
    Sub showList4(ByVal vIDNO As String)
        'Const CST_ERRMSG3 As String = "就服資料庫連線異常!"
        If objconn3 Is Nothing Then [Labmsg4].Text = CST_ERRMSG3 '"就服資料庫連線異常!"
        If objconn3 Is Nothing Then Exit Sub
        Try
            DbAccess.Open(objconn3)
        Catch ex As Exception
            [Labmsg4].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try

        Dim odt As New DataTable
        '(連接 "就服" Oracle DB, 所以不用調整SQL語法)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT j.idno" & vbCrLf
        'sql += "       ,j.name" & vbCrLf
        sql += "        ,j.birth" & vbCrLf
        sql += "        ,f_tech_name(j.techcd) Name" & vbCrLf
        sql += "        ,j.techlv class" & vbCrLf
        sql += "        ,j.tech_date appday" & vbCrLf
        sql += " FROM apltbl_tech j" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += "    AND j.idno = @IDNO" & vbCrLf
        'sql += "   --AND TO_CHAR(j.birth, 'YYYY/MM/DD') = " + comfum.convert_to_dbdate(strBirth) + "" & vbCrLf
        sql += "    AND techcd IS NOT NULL" & vbCrLf
        sql += "    AND techlv IS NOT NULL" & vbCrLf
        sql += "    AND tech_date >= to_date(convert(varchar, add_months(getdate(),-24), 112),'YYYY-MM-DD')" & vbCrLf
        sql += " ORDER BY tech_date DESC" & vbCrLf
        Dim cmd As New SqlCommand(sql, objconn3)

        Try
            With cmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = vIDNO
                odt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            [Labmsg4].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try


        Dim dt As DataTable
        '建立DataGird用的DataTable格式 Start
        dt = New DataTable
        dt.Columns.Add(New DataColumn("seqno"))
        dt.Columns.Add(New DataColumn("Name"))
        dt.Columns.Add(New DataColumn("class"))
        dt.Columns.Add(New DataColumn("appday"))
        '建立DataGird用的DataTable格式 End

        For Each odr As DataRow In odt.Rows
            Dim dr As DataRow = dt.NewRow
            dt.Rows.Add(dr)
            dr("seqno") = dt.Rows.Count
            dr("Name") = odr("Name") ' "程式設計人員"
            dr("class") = ChgTechlv(odr("class")) '"丙級"
            dr("appday") = odr("appday") '"2007/10/02"
        Next

#Region "(No Use)"

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("Name") = "程式設計人員"
        'dr("class") = "丙級"
        'dr("appday") = "2007/10/02"

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("Name") = "會計精算人員"
        'dr("class") = "丙級"
        'dr("appday") = "2008/10/02"

#End Region

        'dt.AcceptChanges()
        [Labmsg4].Text = "查無資料!"
        If dt.Rows.Count > 0 Then
            [Labmsg4].Text = ""
            'dt.DefaultView.Sort = "IDNO,Birthday,TRound"
            DataGrid4.DataSource = dt
            DataGrid4.DataBind()
        End If
    End Sub

    '工作經歷 就服oracle
    Sub showList5(ByVal vIDNO As String)
        If objconn3 Is Nothing Then [Labmsg5].Text = CST_ERRMSG3 '"就服資料庫連線異常!"
        If objconn3 Is Nothing Then Exit Sub
        Try
            DbAccess.Open(objconn3)
        Catch ex As Exception
            [Labmsg5].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try

        'Dim dr As DataRow
        Dim odt As New DataTable
        '(連接 "就服" Oracle DB, 所以不用調整SQL語法)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.idno ,a.name ,a.birth" & vbCrLf
        sql += "        ,j.compname OrgName ,f_business_name(j.profcd) trainName ,f_job_name(j.cjob_no) workName" & vbCrLf
        sql += "        ,j.job_title workName2 ,j.job_desc workMemo ,j.job_salary salary" & vbCrLf
        sql += "        ,to_char(j.job_start_date,'yyyy/mm')+'~'+to_char(j.job_stop_date,'yyyy/mm') workTime" & vbCrLf
        sql += " FROM apltbl_joblist j" & vbCrLf
        sql += " JOIN apltbl a ON a.labor_id = j.labor_id" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND a.idno = @IDNO" & vbCrLf
        sql += " AND j.job_stop_date >= to_date(convert(varchar, add_months(getdate(),-24), 112),'YYYY-MM-DD')" & vbCrLf
        sql += " ORDER BY j.job_start_date DESC" & vbCrLf
        Dim cmd As New SqlCommand(sql, objconn3)

        Try
            With cmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = vIDNO
                odt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            [Labmsg5].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try
        'dt = DbAccess.GetDataTable(sql, objconn)

        '建立DataGird用的DataTable格式 Start
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("seqno"))
        dt.Columns.Add(New DataColumn("OrgName"))
        dt.Columns.Add(New DataColumn("trainName"))
        dt.Columns.Add(New DataColumn("workName"))
        dt.Columns.Add(New DataColumn("workName2"))
        dt.Columns.Add(New DataColumn("workMemo"))
        dt.Columns.Add(New DataColumn("salary"))
        dt.Columns.Add(New DataColumn("workTime"))
        '建立DataGird用的DataTable格式 End

        For Each odr As DataRow In odt.Rows
            Dim dr As DataRow = dt.NewRow
            dt.Rows.Add(dr)
            dr("seqno") = dt.Rows.Count
            dr("OrgName") = odr("OrgName")
            dr("trainName") = odr("trainName") '"其他綜合商品零售業"
            dr("workName") = odr("workName") '"賣場店員及門市人員"
            dr("workName2") = odr("workName2") '"門市人員"
            dr("workMemo") = odr("workMemo") '"銷售"
            dr("salary") = odr("salary") '"26000"
            dr("workTime") = odr("workTime") '"90/05~100/05"
        Next

#Region "(No Use)"

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("OrgName") = "何佳仁書局"
        'dr("trainName") = "書籍、文具零售業"
        'dr("workName") = "賣場店員及門市人員"
        'dr("workName2") = "門市人員"
        'dr("workMemo") = "銷售"
        'dr("salary") = "22000"
        'dr("workTime") = "86/11~90/02"

#End Region

        'dt.AcceptChanges()

        [Labmsg5].Text = "查無資料!"
        If dt.Rows.Count > 0 Then
            [Labmsg5].Text = ""
            'dt.DefaultView.Sort = "IDNO,Birthday,TRound"
            DataGrid5.DataSource = dt
            DataGrid5.DataBind()
        End If
    End Sub

    '推介歷程 就服oracle
    Sub showList6(ByVal vIDNO As String)
        If objconn3 Is Nothing Then [Labmsg6].Text = "就服資料庫連線異常!"
        If objconn3 Is Nothing Then Exit Sub
        Try
            DbAccess.Open(objconn3)
        Catch ex As Exception
            [Labmsg6].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try

        Dim odt As New DataTable
        '(連接 "就服" Oracle DB, 所以不用調整SQL語法)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.idno" & vbCrLf
        sql += "        ,a.name" & vbCrLf
        sql += "        ,a.birth" & vbCrLf
        sql += "        ,j.intr_unkey" & vbCrLf
        sql += "        ,CONVERT(varchar, j.intr_irgdate, 111) interDate1" & vbCrLf
        'sql += "       --,convert(varchar, add_months(getdate(),-24), 112) t2" & vbCrLf
        sql += "        ,c.compname OrgName" & vbCrLf
        'sql += "       --,j.intr_employer_id" & vbCrLf
        'sql += "       --,j.intr_hire_id" & vbCrLf
        sql += "        ,j.intr_occucd Occucd" & vbCrLf
        sql += "        ,j.reply_ansres" & vbCrLf
        'sql += "       ,j.reply_non_fill" & vbCrLf
        'sql += "       ,j.reply_non_fill_other" & vbCrLf
        sql += "        ,j.reply_non_intro" & vbCrLf
        sql += "        ,j.reply_non_intro_other" & vbCrLf
        sql += " FROM intrtbll j" & vbCrLf
        sql += " JOIN apltbl a ON j.intr_labor_id = a.labor_id" & vbCrLf
        sql += " JOIN vnbtbl c ON c.idno = j.intr_idno2" & vbCrLf
        sql += " WHERE j.del_state = 0" & vbCrLf
        sql += "    AND a.IDNO = @IDNO" & vbCrLf
        'sql += "   --AND j.intr_irgdate >= to_date(convert(varchar, add_months(getdate(),-24), 112),'YYYY-MM-DD')" & vbCrLf
        sql += "    AND j.intr_irgdate >= trunc(add_months(getdate(),-24))" & vbCrLf
        'sql += "   AND rownum <= 10" & vbCrLf
        sql += " ORDER BY j.intr_irgdate DESC" & vbCrLf
        Dim cmd As New SqlCommand(sql, objconn3)

        Try
            With cmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = vIDNO
                odt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            [Labmsg6].Text = CST_ERRMSG3
            Common.MessageBox(Me, CST_ERRMSG3)
            Return
        End Try

        '建立DataGird用的DataTable格式 Start
        Dim dt As DataTable
        dt = New DataTable
        dt.Columns.Add(New DataColumn("seqno"))
        dt.Columns.Add(New DataColumn("interDate1"))
        dt.Columns.Add(New DataColumn("OrgName"))
        dt.Columns.Add(New DataColumn("workName"))
        dt.Columns.Add(New DataColumn("workYN"))
        dt.Columns.Add(New DataColumn("NRReason"))
        dt.Columns.Add(New DataColumn("NWReason"))
        '建立DataGird用的DataTable格式 End

        For Each odr As DataRow In odt.Rows
            Dim dr As DataRow = dt.NewRow
            dt.Rows.Add(dr)
            dr("seqno") = dt.Rows.Count
            dr("interDate1") = odr("interDate1") ' "2012/07/18"
            dr("OrgName") = odr("OrgName") '"基隆港務局"
            dr("workName") = funOccucd(odr("Occucd")) '"其他環境清潔維護工及有關體力工"
            dr("workYN") = funShareName(odr("reply_ansres")) '"介紹錄用"
            dr("NRReason") = funReplyNonFill(Convert.ToString(odr("intr_unkey")), Convert.ToString(odr("reply_ansres")), Convert.ToString(odr("reply_non_intro")), Convert.ToString(odr("reply_non_intro_other"))) '""'labReplyAnsres
            dr("NWReason") = funReplyNonIntro(Convert.ToString(odr("intr_unkey")), Convert.ToString(odr("reply_ansres")), Convert.ToString(odr("reply_non_intro")), Convert.ToString(odr("reply_non_intro_other"))) '""'labReplyAnsres2
        Next

#Region "(No Use)"

        '<asp@BoundColumn DataField="seqno" HeaderText="項次"></asp@BoundColumn>
        '                            <asp@BoundColumn DataField="interDate1" HeaderText="介紹日期"></asp@BoundColumn>
        '                            <asp@BoundColumn DataField="orgName" HeaderText="求才單位"></asp@BoundColumn>
        '                            <asp@BoundColumn DataField="workName" HeaderText="職稱名稱"></asp@BoundColumn>
        '                            <asp@BoundColumn DataField="workYN" HeaderText="僱用與否"></asp@BoundColumn>
        '                            <asp@BoundColumn DataField="NRReason" HeaderText="僱主未能錄用原因"></asp@BoundColumn>
        '                            <asp@BoundColumn DataField="NWReason" HeaderText="求職未能推介原因"></asp@BoundColumn>

        'Dim dr As DataRow
        'Dim dt As DataTable
        ''建立DataGird用的DataTable格式 Start
        'dt = New DataTable
        'dt.Columns.Add(New DataColumn("seqno"))
        'dt.Columns.Add(New DataColumn("interDate1"))
        'dt.Columns.Add(New DataColumn("OrgName"))
        'dt.Columns.Add(New DataColumn("workName"))
        'dt.Columns.Add(New DataColumn("workYN"))
        'dt.Columns.Add(New DataColumn("NRReason"))
        'dt.Columns.Add(New DataColumn("NWReason"))
        ''建立DataGird用的DataTable格式 End
        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("interDate1") = "2012/07/18"
        'dr("OrgName") = "基隆港務局"
        'dr("workName") = "其他環境清潔維護工及有關體力工"
        'dr("workYN") = "介紹錄用"
        'dr("NRReason") = ""
        'dr("NWReason") = ""

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("interDate1") = "2012/01/09"
        'dr("OrgName") = "何朝榮診所"
        'dr("workName") = "工讀生"
        'dr("workYN") = "介紹未錄用"
        'dr("NRReason") = "求才廠商已錄用滿額"
        'dr("NWReason") = ""

        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("seqno") = dt.Rows.Count
        'dr("interDate1") = "2012/01/02"
        'dr("OrgName") = "阿瘦實業-基隆店"
        'dr("workName") = "零售業務員"
        'dr("workYN") = "介紹未錄用"
        'dr("NRReason") = "未備駕照或交通工具"
        'dr("NWReason") = ""

#End Region

        dt.AcceptChanges()

        [Labmsg6].Text = "查無資料!"
        If dt.Rows.Count > 0 Then
            [Labmsg6].Text = ""
            'dt.DefaultView.Sort = "IDNO,Birthday,TRound"
            DataGrid6.DataSource = dt
            DataGrid6.DataBind()
        End If
    End Sub

    '技檢級別
    Function ChgTechlv(ByVal Vid As String) As String
        Dim rst As String = ""
        Select Case Vid
            Case "0"
                rst = "(單一)"
            Case "1"
                rst = "(甲級)"
            Case "2"
                rst = "(乙級)"
            Case "3"
                rst = "(丙級)"
        End Select
        Return rst
    End Function
End Class