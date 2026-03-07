Partial Class SD_02_002
    Inherits AuthBasePage

    'INSERT / UPDATE STUD_SELRESULT
    '計算名次(依成績) '有輸入成績
    '系統決定正取、備取、未錄取
    '計算名次(依報名先後) '未輸入成績。
    '系統決定正取、備取、未錄取

    'Protected WithEvents ds As New System.Data.DataSet
    'Protected WithEvents dv1 As New System.Data.DataView
    'Protected WithEvents dv2 As New System.Data.DataView

    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    'Dim FunDr As DataRow
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

        '非 ROLEID=0 LID=0
        'Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
        End If

#Region "(No Use)"

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
        '        FunDr = FunDrArray(0)
        '        If RIDValue.Value <> sm.UserInfo.RID Then
        '            Button2.Enabled = False
        '            Button3.Enabled = False
        '            TIMS.Tooltip(Button2, "登入權限與計劃權限不符，停用止功能", True)
        '            TIMS.Tooltip(Button3, "登入權限與計劃權限不符，停用止功能", True)
        '        Else
        '            Button2.Enabled = True
        '            Button3.Enabled = True
        '            If FunDr("Adds") = 1 Then
        '                Button2.Enabled = True
        '                Button3.Enabled = True
        '            Else
        '                Button2.Enabled = False
        '                Button3.Enabled = False
        '                TIMS.Tooltip(Button2, "無新增功能權限", True)
        '                TIMS.Tooltip(Button3, "無新增功能權限", True)
        '            End If
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            Button1.Enabled = True
        '        Else
        '            Button1.Enabled = False
        '            TIMS.Tooltip(Button1, "無查詢功能權限", True)
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

#End Region

        Button2.Enabled = True
        Button3.Enabled = True
        If RIDValue.Value <> sm.UserInfo.RID Then
            Button2.Enabled = False
            Button3.Enabled = False
            TIMS.Tooltip(Button2, "登入權限與計劃權限不符，停用止功能", True)
            TIMS.Tooltip(Button3, "登入權限與計劃權限不符，停用止功能", True)
        End If

        Button1.Enabled = True

        If Not IsPostBack Then
            msg1.Text = ""
            msg2.Text = ""
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                'Button4_Click(sender, e)
                Call get_onlyone1()
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button1.Attributes("onclick") = "javascript:return chk();"

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

        Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試
        If flag_test Then '測試用
            Button2.Enabled = True '測試用
            Button3.Enabled = True '測試用
        End If '測試用
    End Sub

    '查詢
    Sub search1()
        Table4.Visible = True
        If start_date.Text <> "" Then start_date.Text = Common.FormatDate(start_date.Text)
        If end_date.Text <> "" Then end_date.Text = Common.FormatDate(end_date.Text)
        If classname.Text <> "" Then classname.Text = Trim(classname.Text)

        'Dim Sql As String = ""
        Dim param As SqlParameter
        Dim da As SqlDataAdapter = Nothing
        'Dim DateStr As String = ""
        'Dim cmd As SqlCommand

        Dim DateStr As String = ""
        DateStr = ""
        If start_date.Text <> "" Then DateStr += " AND cc.STDate >= @start_date " & vbCrLf
        If end_date.Text <> "" Then DateStr += " AND cc.STDate <= @end_date " & vbCrLf
        If OCIDValue1.Value <> "" Then DateStr += " AND cc.OCID = @OCID" & vbCrLf
        If cjobValue.Value <> "" Then DateStr += " AND cc.CJOB_UNKEY = @CJOB_UNKEY " & vbCrLf
        If classname.Text <> "" Then DateStr += " AND cc.ClassCName LIKE @ClassCName " & vbCrLf '必填因為要搜尋所有班級 AMU

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WT1 AS (SELECT planid FROM ID_PLAN WHERE TPlanID = '" & sm.UserInfo.TPlanID & "' AND Years = '" & sm.UserInfo.Years & "') " & vbCrLf
        sql += " SELECT cc.ocid " & vbCrLf
        sql += "        ,cc.stdate " & vbCrLf
        sql += "        ,cc.ftdate " & vbCrLf
        sql += "        ,cc.classcname " & vbCrLf
        sql += "        ,cc.cycltype " & vbCrLf
        sql += "        ,cc.leveltype " & vbCrLf
        sql &= "        ,cc.IsCalculate " & vbCrLf
        sql &= "        ,d.ClassID " & vbCrLf
        sql += "        ,b.OCID1 " & vbCrLf
        sql += "        ,b.total " & vbCrLf
        sql += "        ,CASE WHEN b.total >0 THEN 1 END t1 " & vbCrLf
        sql += "        ,CASE WHEN ISNULL(b.total,0) <=0 THEN 1 END t2 " & vbCrLf
        sql += " FROM Class_ClassInfo cc " & vbCrLf
        sql += " JOIN ID_Class d ON cc.CLSID = d.CLSID " & vbCrLf
        sql += " JOIN WT1 ip ON ip.planid = cc.planid " & vbCrLf
        sql += " JOIN (" & vbCrLf
        sql += "   SELECT b.OCID1 ,sum(b.TotalResult) total " & vbCrLf
        sql += "   FROM Stud_EnterType b " & vbCrLf
        sql += "   JOIN Class_ClassInfo cc ON cc.ocid = b.OCID1 " & vbCrLf
        sql += "   JOIN WT1 ip ON ip.planid = cc.planid " & vbCrLf
        sql += "   WHERE 1=1 " & vbCrLf
        sql += "      AND cc.NotOpen = 'N' " & vbCrLf
        If Not flgROLEIDx0xLIDx0 Then sql += "   AND cc.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        sql += "   AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
        sql &= DateStr
        sql += "   GROUP BY b.OCID1 " & vbCrLf
        sql += " ) b ON b.ocid1 = cc.ocid " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        sql += "    AND cc.NotOpen = 'N' " & vbCrLf
        If Not flgROLEIDx0xLIDx0 Then sql += " AND cc.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        sql += " AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
        sql &= DateStr
        'sql += " AND b.total > 0 " & vbCrLf
        sql += " ORDER BY d.ClassID ,cc.CyclType " & vbCrLf

        da = New SqlDataAdapter(sql, objconn)
        If start_date.Text <> "" Then
            param = da.SelectCommand.Parameters.Add("@start_date", SqlDbType.DateTime)
            param.Value = start_date.Text
        End If
        If end_date.Text <> "" Then
            param = da.SelectCommand.Parameters.Add("@end_date", SqlDbType.DateTime)
            param.Value = end_date.Text
        End If
        If OCIDValue1.Value <> "" Then
            param = da.SelectCommand.Parameters.Add("@OCID", SqlDbType.VarChar)
            param.Value = OCIDValue1.Value.ToString
        End If
        If cjobValue.Value <> "" Then
            param = da.SelectCommand.Parameters.Add("@CJOB_UNKEY", SqlDbType.VarChar)
            param.Value = cjobValue.Value.ToString
        End If
        If classname.Text <> "" Then
            param = da.SelectCommand.Parameters.Add("@ClassCName", SqlDbType.NVarChar, 50)
            param.Value = "%" & classname.Text & "%"
        End If
        Dim ds As New System.Data.DataSet
        da.Fill(ds, "search1")

        Dim ff As String = ""
        ff = "t1=1"
        ds.Tables("search1").DefaultView.RowFilter = ff
        Dim dt1 As DataTable = TIMS.dv2dt(ds.Tables("search1").DefaultView)
        With DataGrid1
            .DataSource = dt1
            .DataBind()
        End With
        ff = "t2=1"
        ds.Tables("search1").DefaultView.RowFilter = ff
        Dim dt2 As DataTable = TIMS.dv2dt(ds.Tables("search1").DefaultView)
        With DataGrid2
            .DataSource = dt2
            .DataBind()
        End With

#Region "(No Use)"

        'sql = "" & vbCrLf
        'sql += " select cc.ocid" & vbCrLf
        'sql += " ,cc.stdate" & vbCrLf
        'sql += " ,cc.ftdate" & vbCrLf
        'sql += " ,cc.classcname" & vbCrLf
        'sql += " ,cc.cycltype" & vbCrLf
        'sql += " ,cc.leveltype" & vbCrLf
        'sql &= " ,cc.IsCalculate" & vbCrLf
        'sql &= " ,d.ClassID" & vbCrLf
        'sql += " ,b.OCID1" & vbCrLf
        'sql += " ,b.total" & vbCrLf
        'sql += " from Class_ClassInfo cc" & vbCrLf
        'sql += " join ID_Class d on cc.CLSID=d.CLSID" & vbCrLf
        'sql += " join (" & vbCrLf
        'sql += "   select b.OCID1,sum(b.TotalResult) total " & vbCrLf
        'sql += "   from Stud_EnterType b" & vbCrLf
        'sql += "   join Class_ClassInfo cc on cc.ocid =b.ocid1" & vbCrLf
        'sql += "   where 1=1" & vbCrLf
        'sql += "   and cc.NotOpen='N' " & vbCrLf
        'If Not flgROLEIDx0xLIDx0 Then
        '    sql += "   and cc.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
        'End If
        'sql += "   and cc.RID='" & RIDValue.Value & "'" & vbCrLf
        'sql &= DateStr
        'sql += "   group by b.OCID1 " & vbCrLf
        'sql += " ) b on b.ocid1 =cc.ocid" & vbCrLf
        'sql += " where 1=1" & vbCrLf
        'sql += " and cc.NotOpen='N' " & vbCrLf
        'If Not flgROLEIDx0xLIDx0 Then
        '    sql += " and cc.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
        'End If
        'sql += " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
        'sql &= DateStr
        'sql += " and dbo.NVL(b.total,0)<=0" & vbCrLf
        'sql += " Order By d.ClassID,cc.CyclType" & vbCrLf

        'da = New SqlDataAdapter(sql, objconn)
        'If start_date.Text <> "" Then
        '    param = da.SelectCommand.Parameters.Add("@start_date", SqlDbType.DateTime)
        '    param.Value = CDate(start_date.Text)
        'End If
        'If end_date.Text <> "" Then
        '    param = da.SelectCommand.Parameters.Add("@end_date", SqlDbType.DateTime)
        '    param.Value = CDate(end_date.Text)
        'End If
        'If OCIDValue1.Value <> "" Then
        '    param = da.SelectCommand.Parameters.Add("@OCID", SqlDbType.VarChar)
        '    param.Value = OCIDValue1.Value.ToString
        'End If
        'If cjobValue.Value <> "" Then
        '    param = da.SelectCommand.Parameters.Add("@CJOB_UNKEY", SqlDbType.VarChar)
        '    param.Value = cjobValue.Value.ToString
        'End If
        'If classname.Text <> "" Then
        '    param = da.SelectCommand.Parameters.Add("@ClassCName", SqlDbType.NVarChar, 50)
        '    param.Value = "%" & classname.Text & "%"
        'End If
        'param = da.SelectCommand.Parameters.Add("@classID", SqlDbType.VarChar, 4)
        'If classID.Text = "" Then
        '    param.Value = "%"
        'Else
        '    param.Value = classID.Text
        'End If
        'If OCIDValue1.Value = "" Then
        '    param.Value = "%"
        'Else
        '    param.Value = OCIDValue1.Value.ToString
        'End If
        'da.Fill(ds, "search2")
        'dv2.Table = ds.Tables("search2")

#End Region

        msg1.Visible = True
        msg1.Text = ""
        If dt1.Rows.Count = 0 Then
            DataGrid1.Visible = False
            Button2.Visible = False
            'msg1.Visible = True
            msg1.Text = "查無資料!!"
            TIMS.Tooltip(msg1, "若有開班，但查無資料，是沒有報名資料造成。")
        Else
            Button2.Visible = True
            DataGrid1.Visible = True
            'msg1.Visible = False
            DataGrid1.DataSource = dt1
            DataGrid1.DataBind()
        End If

        msg2.Visible = True
        msg2.Text = ""
        If dt2.Rows.Count = 0 Then
            DataGrid2.Visible = False
            Button3.Visible = False
            'msg2.Visible = True
            msg2.Text = "查無資料!!"
            TIMS.Tooltip(msg2, "若有開班，但查無資料，是沒有報名資料造成。")
        Else
            Button3.Visible = True
            DataGrid2.Visible = True
            'msg2.Visible = False
            DataGrid2.DataSource = dt2
            DataGrid2.DataBind()
        End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call search1()
    End Sub

    Function SaveData1(ByVal vOCID As String) As String
        Dim msgbox As String = "" 'err msg
        '2006/03/ add conn by matt
        'Dim daResult As SqlDataAdapter = Nothing
        Dim dtResult As DataTable = Nothing 'STUD_SELRESULT
        Dim sql As String = ""
        sql = " SELECT * FROM STUD_SELRESULT WHERE OCID = '" & vOCID & "' "
        dtResult = DbAccess.GetDataTable(sql, objconn)

        '取出授課人數
        Dim iTnum As Integer = 0
        Dim ClassCName As String = ""
        Dim STDate As Date '開訓日期
        Dim drC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If Not drC Is Nothing Then
            iTnum = drC("Tnum")
            ClassCName = drC("ClassCName2")
            STDate = CDate(drC("STDate")) '開訓日期
        End If

        '取出選擇第一志願班級的學生
        Dim dt1 As DataTable = Nothing 'STUD_ENTERTYPE 
        sql = " SELECT * FROM STUD_ENTERTYPE WHERE OCID1 = '" & vOCID & "' AND CCLID IS NULL "
        dt1 = DbAccess.GetDataTable(sql, objconn)

        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            Dim da As SqlDataAdapter = Nothing
            Dim dt As DataTable = Nothing
            sql = " SELECT * FROM STUD_SELRESULT WHERE OCID = '" & vOCID & "' "
            dt = DbAccess.GetDataTable(sql, da, objTrans)

            '取出不用考試的學生
            For Each dr1 As DataRow In dt1.Select("NotExam=1")
                Dim dr As DataRow = Nothing
                If dt.Select("SETID='" & dr1("SETID") & "' AND EnterDate='" & dr1("EnterDate") & "' AND SerNum='" & dr1("SerNum") & "'").Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("SETID") = dr1("SETID")
                    dr("EnterDate") = dr1("EnterDate")
                    dr("SerNum") = dr1("SerNum")
                Else
                    dr = dt.Select("SETID='" & dr1("SETID") & "' AND EnterDate='" & dr1("EnterDate") & "' AND SerNum='" & dr1("SerNum") & "'")(0)
                End If
                dr("OCID") = dr1("OCID1")
                dr("SumOfGrad") = dr1("TotalResult")
                dr("TRNDType") = dr1("TRNDType")
                'If RIDValue.Value <> "" Then dr("RID") = RIDValue.Value Else dr("RID") = sm.UserInfo.RID
                dr("RID") = sm.UserInfo.RID
                dr("PlanID") = sm.UserInfo.PlanID
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                If STDate >= Now Then  '假如是開訓前
                    If iTnum > 0 Then
                        dr("SelResultID") = TIMS.cst_SelResultID_正取 ' "01"
                        iTnum -= 1
                    Else
                        'dr("SelResultID") = "02"
                        dr("SelResultID") = TIMS.cst_SelResultID_未錄取 '"03"
                    End If
                Else                    '假如是開訓後
                    If iTnum > 0 Then
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case TIMS.cst_SelResultID_正取 '"01" '正取
                                iTnum -= 1
                            Case TIMS.cst_SelResultID_備取 '"02" '備取
                            Case TIMS.cst_SelResultID_未錄取 '"03" '未錄取
                            Case Else
                                dr("SelResultID") = TIMS.cst_SelResultID_正取 '"01" '正取
                                iTnum -= 1
                        End Select
                    Else
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case TIMS.cst_SelResultID_正取 '"01" '正取
                                iTnum -= 1
                            Case TIMS.cst_SelResultID_備取 '"02" '備取
                            Case TIMS.cst_SelResultID_未錄取 '"03" '未錄取
                            Case Else
                                dr("SelResultID") = TIMS.cst_SelResultID_未錄取 '"03" '未錄取
                        End Select
                    End If
                End If
            Next

            '取出甲券學生
            Dim ff As String = ""
            For Each dr1 As DataRow In dt1.Select("NotExam=0 and TRNDType=1", "TotalResult Desc")
                ff = "SETID='" & dr1("SETID") & "' AND EnterDate='" & dr1("EnterDate") & "' AND SerNum='" & dr1("SerNum") & "'"
                Dim dr As DataRow = Nothing
                If dt.Select(ff).Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("SETID") = dr1("SETID")
                    dr("EnterDate") = dr1("EnterDate")
                    dr("SerNum") = dr1("SerNum")
                Else
                    dr = dt.Select(ff)(0)
                End If
                dr("OCID") = dr1("OCID1")
                dr("SumOfGrad") = dr1("TotalResult")
                dr("TRNDType") = dr1("TRNDType")
                'If RIDValue.Value <> "" Then dr("RID") = RIDValue.Value Else dr("RID") = sm.UserInfo.RID
                dr("RID") = sm.UserInfo.RID
                dr("PlanID") = sm.UserInfo.PlanID
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                If STDate >= Now Then  '假如是開訓前
                    If iTnum > 0 Then
                        dr("SelResultID") = TIMS.cst_SelResultID_正取 '"01"
                        iTnum -= 1
                    Else
                        'dr("SelResultID") = "02"
                        dr("SelResultID") = TIMS.cst_SelResultID_未錄取 '"03"
                    End If
                Else                    '假如是開訓後
                    If iTnum > 0 Then
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case TIMS.cst_SelResultID_正取 '"01" '正取
                                iTnum -= 1
                            Case TIMS.cst_SelResultID_備取 '"02" '備取
                            Case TIMS.cst_SelResultID_未錄取 '"03" '未錄取
                            Case Else
                                dr("SelResultID") = "01" '正取
                                iTnum -= 1
                        End Select
                    Else
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01" '正取
                                iTnum -= 1
                            Case "02" '備取
                            Case "03" '未錄取
                            Case Else
                                dr("SelResultID") = "03" '未錄取
                        End Select
                    End If
                End If
            Next

            '取出剩下的學生
            For Each dr1 As DataRow In dt1.Select("NotExam=0 and (TRNDType<>1 or TRNDType IS NULL)", "TotalResult Desc")
                ff = "SETID='" & dr1("SETID") & "' and EnterDate='" & dr1("EnterDate") & "' and SerNum='" & dr1("SerNum") & "'"
                Dim dr As DataRow = Nothing
                If dt.Select(ff).Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("SETID") = dr1("SETID")
                    dr("EnterDate") = dr1("EnterDate")
                    dr("SerNum") = dr1("SerNum")
                Else
                    dr = dt.Select(ff)(0)
                End If
                dr("OCID") = dr1("OCID1")
                dr("SumOfGrad") = dr1("TotalResult")
                dr("TRNDType") = dr1("TRNDType")
                'If RIDValue.Value <> "" Then dr("RID") = RIDValue.Value Else dr("RID") = sm.UserInfo.RID
                dr("RID") = sm.UserInfo.RID
                dr("PlanID") = sm.UserInfo.PlanID
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                If STDate >= Now Then  '假如是開訓前
                    If iTnum > 0 Then
                        dr("SelResultID") = "01"
                        iTnum -= 1
                    Else
                        'dr("SelResultID") = "02"
                        dr("SelResultID") = "03"
                    End If
                Else                    '假如是開訓後
                    If iTnum > 0 Then
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01" '正取
                                iTnum -= 1
                            Case "02" '備取
                            Case "03" '未錄取
                            Case Else
                                dr("SelResultID") = "01" '正取
                                iTnum -= 1
                        End Select
                    Else
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01" '正取
                                iTnum -= 1
                            Case "02" '備取
                            Case "03" '未錄取
                            Case Else
                                dr("SelResultID") = "03" '未錄取
                        End Select
                    End If
                End If
            Next
            DbAccess.UpdateDataTable(dt, da, objTrans)

            sql = " UPDATE CLASS_CLASSINFO SET IsCalculate = 'Y' WHERE OCID = '" & vOCID & "' "
            DbAccess.ExecuteNonQuery(sql, objTrans)

            DbAccess.CommitTrans(objTrans)
        Catch ex As Exception
            Dim sMailMsg As String = ""
            sMailMsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(sMailMsg, ex)
            DbAccess.RollbackTrans(objTrans)
            msgbox += "試算失敗!失敗班級:" & ClassCName & vbCrLf
        End Try
        Call TIMS.CloseDbConn(tConn)
        Return msgbox
    End Function

    Function SaveData2(ByVal vOCID As String) As String
        Dim msgbox As String = ""

        '取出授課人數
        Dim iTnum As Integer = 0
        Dim ClassCName As String = ""
        Dim STDate As Date '開訓日期
        Dim drC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If Not drC Is Nothing Then
            iTnum = drC("Tnum")
            ClassCName = drC("ClassCName2")
            STDate = CDate(drC("STDate")) '開訓日期
        End If

        Dim sql As String = ""

        '取出選擇第一志願班級的學生
        Dim dtE1 As DataTable = Nothing 'STUD_ENTERTYPE 
        sql = " SELECT * FROM STUD_ENTERTYPE WHERE OCID1 = '" & vOCID & "' AND CCLID IS NULL "
        dtE1 = DbAccess.GetDataTable(sql, objconn)

        'Dim daResult As SqlDataAdapter = Nothing
        Dim dtResult As DataTable = Nothing 'STUD_SELRESULT
        sql = " SELECT * FROM STUD_SELRESULT WHERE OCID = '" & vOCID & "' "
        dtResult = DbAccess.GetDataTable(sql, objconn)

#Region "(No Use)"

        'Sql = "select TNum,ClassCName,CyclType,LevelType,STDate from Class_ClassInfo where OCID='" & vOCID & "'"
        'da = New SqlDataAdapter(Sql, objconn)
        'da.Fill(ds, "Tnum")
        'If ds.Tables("Tnum").Rows.Count <> 0 Then
        '    Tnum = ds.Tables("Tnum").Rows(0).Item("Tnum")
        '    ClassCName = ds.Tables("Tnum").Rows(0).Item("ClassCName") & "第" & TIMS.GetChtNum(CInt(ds.Tables("Tnum").Rows(0).Item("CyclType"))) & "期"
        '    STDate = CDate(ds.Tables("Tnum").Rows(0).Item("STDate")) '開訓日期
        'End If

        'Dim ds As New System.Data.DataSet
        'Dim dtE1 As New DataTable
        'Sql = "select * from Stud_EnterType where OCID1='" & vOCID & "' and CCLID IS NULL"
        'da = New SqlDataAdapter(Sql, objconn)
        'da.Fill(dtE1)

#End Region

        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            Dim da As SqlDataAdapter = Nothing
            Dim dt As DataTable = Nothing
            sql = " SELECT * FROM STUD_SELRESULT WHERE OCID = '" & vOCID & "' "
            dt = DbAccess.GetDataTable(sql, da, objTrans)

            Dim row() As DataRow
            row = dtE1.Select("TRNDType=1", "RelEnterDate")
            For j As Integer = 0 To row.Length - 1
                Dim dr As DataRow = Nothing
                If dt.Select("SETID='" & row(j)("SETID") & "' and EnterDate='" & row(j)("EnterDate") & "' and SerNum='" & row(j)("SerNum") & "'").Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("SETID") = row(j)("SETID")
                    dr("EnterDate") = row(j)("EnterDate")
                    dr("SerNum") = row(j)("SerNum")
                Else
                    dr = dt.Select("SETID='" & row(j)("SETID") & "' and EnterDate='" & row(j)("EnterDate") & "' and SerNum='" & row(j)("SerNum") & "'")(0)
                End If

                dr("OCID") = row(j)("OCID1")
                dr("SumOfGrad") = row(j)("TotalResult")
                dr("TRNDType") = row(j)("TRNDType")
                'If RIDValue.Value <> "" Then dr("RID") = RIDValue.Value Else dr("RID") = sm.UserInfo.RID
                dr("RID") = sm.UserInfo.RID
                dr("PlanID") = sm.UserInfo.PlanID
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                If STDate >= Now Then  '假如是開訓前
                    If iTnum > 0 Then
                        dr("SelResultID") = "01" '錄取
                        iTnum -= 1
                    Else
                        'dr("SelResultID") = "02"
                        dr("SelResultID") = "03" '未錄取
                    End If
                Else                    '假如是開訓後
                    If iTnum > 0 Then
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01" '正取
                                iTnum -= 1
                            Case "02" '備取
                            Case "03" '未錄取
                            Case Else
                                dr("SelResultID") = "01" '正取
                                iTnum -= 1
                        End Select
                    Else
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01" '正取
                                iTnum -= 1
                            Case "02" '備取
                            Case "03" '未錄取
                            Case Else
                                dr("SelResultID") = "03" '未錄取
                        End Select
                    End If
                End If
            Next

            row = dtE1.Select("TRNDType<>1 OR TRNDType IS NULL", "RelEnterDate")
            For j As Integer = 0 To row.Length - 1
                Dim dr As DataRow = Nothing
                If dt.Select("SETID='" & row(j)("SETID") & "' AND EnterDate='" & row(j)("EnterDate") & "' AND SerNum='" & row(j)("SerNum") & "'").Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("SETID") = row(j)("SETID")
                    dr("EnterDate") = row(j)("EnterDate")
                    dr("SerNum") = row(j)("SerNum")
                Else
                    dr = dt.Select("SETID='" & row(j)("SETID") & "' AND EnterDate='" & row(j)("EnterDate") & "' AND SerNum='" & row(j)("SerNum") & "'")(0)
                End If
                dr("OCID") = row(j)("OCID1")
                dr("SumOfGrad") = row(j)("TotalResult")
                dr("TRNDType") = row(j)("TRNDType")
                'If RIDValue.Value <> "" Then dr("RID") = RIDValue.Value Else dr("RID") = sm.UserInfo.RID
                dr("RID") = sm.UserInfo.RID
                dr("PlanID") = sm.UserInfo.PlanID
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                If STDate >= Now Then  '假如是開訓前
                    If iTnum > 0 Then
                        dr("SelResultID") = "01"
                        iTnum -= 1
                    Else
                        'dr("SelResultID") = "02"
                        dr("SelResultID") = "03"
                    End If
                Else                    '假如是開訓後
                    If iTnum > 0 Then
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01" '正取
                                iTnum -= 1
                            Case "02" '備取
                            Case "03" '未錄取
                            Case Else
                                dr("SelResultID") = "01" '正取
                                iTnum -= 1
                        End Select
                    Else
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01" '正取
                                iTnum -= 1
                            Case "02" '備取
                            Case "03" '未錄取
                            Case Else
                                dr("SelResultID") = "03" '未錄取
                        End Select
                    End If
                End If
            Next
            DbAccess.UpdateDataTable(dt, da, objTrans)

            sql = " UPDATE CLASS_CLASSINFO SET IsCalculate = 'Y' WHERE OCID = '" & vOCID & "' "
            DbAccess.ExecuteNonQuery(sql, objTrans)

            DbAccess.CommitTrans(objTrans)
        Catch ex As Exception
            Dim sMailMsg As String = ""
            sMailMsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(sMailMsg, ex)
            DbAccess.RollbackTrans(objTrans)
            msgbox += "試算失敗!失敗班級:" & ClassCName & vbCrLf
        End Try
        Call TIMS.CloseDbConn(tConn)

        Return msgbox
    End Function

    '計算名次(依成績)
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim rqOCID_Grade As String = Request("OCID_Grade")
        rqOCID_Grade = TIMS.ClearSQM(rqOCID_Grade)
        If rqOCID_Grade = "" Then
            Common.MessageBox(Me, "請勾選班級!")
            Exit Sub
        End If
        If rqOCID_Grade <> "" Then
            rqOCID_Grade = TIMS.CombiSQM2IN(rqOCID_Grade)
            rqOCID_Grade = Replace(rqOCID_Grade, "'", "")
        End If
        If rqOCID_Grade = "" Then
            Common.MessageBox(Me, "請勾選班級!")
            Exit Sub
        End If

        Dim msgbox As String = "" 'err msg
        Dim strAll As String() = Split(rqOCID_Grade, ",")
        For i As Integer = 0 To strAll.Length - 1
            strAll(i) = TIMS.ClearSQM(strAll(i))
            If strAll(i) <> "" Then
                msgbox &= SaveData1(strAll(i))
            End If
        Next
        If msgbox = "" Then
            Common.MessageBox(Me, "資料試算成功!")
        Else
            Common.MessageBox(Me, msgbox)
        End If

        Call search1()
        'Button1_Click(sender, e)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            If Int(drv("CyclType")) <> 0 Then e.Item.Cells(1).Text += "第" & Int(drv("CyclType")) & "期"
            If Not IsDBNull(drv("LevelType")) Then
                If Int(drv("LevelType")) <> 0 Then e.Item.Cells(1).Text += "第" & Int(drv("LevelType")) & "期"
            End If
            If drv("IsCalculate") = "Y" Then e.Item.Cells(1).Text += "(總分已計算)"
        End If
    End Sub

    '計算名次(依報名先後)
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim rqOCID_Sort As String = Request("OCID_Sort")
        rqOCID_Sort = TIMS.ClearSQM(rqOCID_Sort)
        If rqOCID_Sort = "" Then
            Common.MessageBox(Me, "請勾選班級!")
            Exit Sub
        End If
        If rqOCID_Sort <> "" Then
            rqOCID_Sort = TIMS.CombiSQM2IN(rqOCID_Sort)
            rqOCID_Sort = Replace(rqOCID_Sort, "'", "")
        End If
        If rqOCID_Sort = "" Then
            Common.MessageBox(Me, "請勾選班級!")
            Exit Sub
        End If

        Dim msgbox As String = "" 'err msg
        Dim strAll As String() = Split(rqOCID_Sort, ",")  '班級代號 陣列
        For i As Integer = 0 To strAll.Length - 1
            strAll(i) = TIMS.ClearSQM(strAll(i))
            If strAll(i) <> "" Then msgbox &= SaveData2(strAll(i))
        Next

        If msgbox = "" Then
            Common.MessageBox(Me, "資料試算成功!")
        Else
            Common.MessageBox(Me, msgbox)
        End If

        Call search1()
        'Button1_Click(sender, e)
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            If Int(drv("CyclType")) <> 0 Then e.Item.Cells(1).Text += "第" & Int(drv("CyclType")) & "期"
            If Not IsDBNull(drv("LevelType")) Then
                If Int(drv("LevelType")) <> 0 Then e.Item.Cells(1).Text += "第" & Int(drv("LevelType")) & "期"
            End If
            If drv("IsCalculate") = "Y" Then e.Item.Cells(1).Text += "(總分已計算)"
        End If
    End Sub

    '判斷機構是否只有一個班級
    Sub get_onlyone1()
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)  '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        'Table4.Style("display") = "none"
        Table4.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        'Table4.Style("display") = "none"
        Table4.Visible = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call get_onlyone1()
    End Sub

    '清除試算
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Common.MessageBox(Me, "該功能並無設計，尚未完成!!")
    End Sub

    '清除試算
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Common.MessageBox(Me, "該功能並無設計，尚未完成!!")
    End Sub
End Class