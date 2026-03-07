Partial Class CP_01_001_add
    Inherits AuthBasePage

    Const cst_SearchStr As String = "SearchStr"
    Const cst_SearchStr2 As String = "_SearchStr" '記錄用
    'Dim FunDr As DataRow
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
        If Not Session(cst_SearchStr) Is Nothing Then Session(cst_SearchStr) = Session(cst_SearchStr) 'Else Session(cst_SearchStr) = Me.ViewState("SearchStr")

        If Not IsPostBack Then
            If Not Session(cst_SearchStr) Is Nothing Then
                center.Text = TIMS.GetMyValue(Session(cst_SearchStr), "center")
                RIDValue.Value = TIMS.GetMyValue(Session(cst_SearchStr), "RIDValue")
                TMID1.Text = TIMS.GetMyValue(Session(cst_SearchStr), "TMID1")
                OCID1.Text = TIMS.GetMyValue(Session(cst_SearchStr), "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(Session(cst_SearchStr), "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(Session(cst_SearchStr), "OCIDValue1")
                EndDate.Value = TIMS.GetMyValue(Session(cst_SearchStr), "end_date")
                ' Session(cst_SearchStr) 不清空返回使用。
            End If
            If Not Session(cst_SearchStr2) Is Nothing Then
                Me.ViewState("_SearchStr") = Session(cst_SearchStr2)
                Session(cst_SearchStr2) = Nothing
            End If
        End If
        Button5.Style("display") = "none"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

        Button1.Attributes("onclick") = "javascript:return chkdata()"
        If Not IsPostBack Then
            If Request("OCID") <> "" Then
                create(Request("OCID"), Request("SeqNo"))
            End If
            If Request("DOCID") <> "" Then
                create(Request("DOCID"), "")
            End If
            FindRange()
        End If

        If Request("view") = "1" Then
            Button1.Visible = False
        End If

#Region "(No Use)"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        'End If
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
        '        If FunDr("Adds") = 1 Then
        '            Button1.Enabled = True
        '        Else
        '            Button1.Enabled = False
        '        End If
        '    End If
        'End If

#End Region

        '檢查帳號的功能權限-----------------------------------End
        'Data5.ClientID()
        Me.TurnoutCount.Attributes("onBlur") = "return Turnoutchk();"
        Me.RejectCount.Attributes("onBlur") = "return Rejectchk();"
    End Sub

    Sub create(ByVal OCID As String, ByVal SeqNo As String)
        Dim sql As String
        sql = " SELECT * FROM Class_Visitor WHERE OCID = '" & OCID & "' AND SeqNo = '" & SeqNo & "' "
        Dim i As Integer = 0
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            ApplyDate.Text = dr("ApplyDate")
            AuthCount.Text = Convert.ToString(dr("AuthCount"))
            TurthCount.Text = Convert.ToString(dr("TurthCount"))
            TurnoutCount.Text = Convert.ToString(dr("TurnoutCount"))
            TruancyCount.Text = Convert.ToString(dr("TruancyCount"))
            RejectCount.Text = Convert.ToString(dr("RejectCount"))
            For i = 0 To Data1.Items.Count - 1
                If Convert.ToString(dr("Data1")) = Data1.Items(i).Value Then Data1.Items(i).Selected = True
            Next
            DataCopy1.Text = Convert.ToString(dr("DataCopy1"))
            Data1Note.Text = Convert.ToString(dr("Data1Note"))
            For i = 0 To Data2.Items.Count - 1
                If Convert.ToString(dr("Data2")) = Data2.Items(i).Value Then Data2.Items(i).Selected = True
            Next
            DataCopy2.Text = Convert.ToString(dr("DataCopy2"))
            Data2Note.Text = Convert.ToString(dr("Data2Note"))
            For i = 0 To Data3.Items.Count - 1
                If Convert.ToString(dr("Data3")) = Data3.Items(i).Value Then Data3.Items(i).Selected = True
            Next
            DataCopy3.Text = Convert.ToString(dr("DataCopy3"))
            Data3Note.Text = Convert.ToString(dr("Data3Note"))
            For i = 0 To Data4.Items.Count - 1
                If Convert.ToString(dr("Data4")) = Data4.Items(i).Value Then Data4.Items(i).Selected = True
            Next
            DataCopy4.Text = Convert.ToString(dr("DataCopy4"))
            Data4Note.Text = Convert.ToString(dr("Data4Note"))
            For i = 0 To Data5.Items.Count - 1
                If Convert.ToString(dr("Data5")) = Data5.Items(i).Value Then Data5.Items(i).Selected = True
            Next
            DataCopy5.Text = Convert.ToString(dr("DataCopy5"))
            Data5Note.Text = Convert.ToString(dr("Data5Note"))
            For i = 0 To Data6.Items.Count - 1
                If Convert.ToString(dr("Data6")) = Data6.Items(i).Value Then Data6.Items(i).Selected = True
            Next
            DataCopy6.Text = Convert.ToString(dr("DataCopy6"))
            Data6Note.Text = Convert.ToString(dr("Data6Note"))
            For i = 0 To Data7.Items.Count - 1
                If Convert.ToString(dr("Data7")) = Data7.Items(i).Value Then Data7.Items(i).Selected = True
            Next
            DataCopy7.Text = Convert.ToString(dr("DataCopy7"))
            Data7Note.Text = Convert.ToString(dr("Data7Note"))
            For i = 0 To Item1_1.Items.Count - 1
                If Convert.ToString(dr("Item1_1")) = Item1_1.Items(i).Value Then Item1_1.Items(i).Selected = True
            Next
            For i = 0 To Item1_2.Items.Count - 1
                If Convert.ToString(dr("Item1_2")) = Item1_2.Items(i).Value Then Item1_2.Items(i).Selected = True
            Next
            Item1_3.Text = Convert.ToString(dr("Item1_3"))
            Item1Pros.Text = Convert.ToString(dr("Item1Pros"))
            Item1Note.Text = Convert.ToString(dr("Item1Note"))
            For i = 0 To Item2_1.Items.Count - 1
                If Convert.ToString(dr("Item2_1")) = Item2_1.Items(i).Value Then Item2_1.Items(i).Selected = True
            Next
            For i = 0 To Item2_2.Items.Count - 1
                If Convert.ToString(dr("Item2_2")) = Item2_2.Items(i).Value Then Item2_2.Items(i).Selected = True
            Next
            Item2Pros.Text = Convert.ToString(dr("Item2Pros"))
            Item2Note.Text = Convert.ToString(dr("Item2Note"))
            For i = 0 To Item3_1.Items.Count - 1
                If Convert.ToString(dr("Item3_1")) = Item3_1.Items(i).Value Then Item3_1.Items(i).Selected = True
            Next
            Item3_1Tech.Text = Convert.ToString(dr("Item3_1Tech"))
            Item3_1Tutor.Text = Convert.ToString(dr("Item3_1Tutor"))
            For i = 0 To Item3_2.Items.Count - 1
                If Convert.ToString(dr("Item3_2")) = Item3_2.Items(i).Value Then Item3_2.Items(i).Selected = True
            Next
            Item3Pros.Text = Convert.ToString(dr("Item3Pros"))
            Item3Note.Text = Convert.ToString(dr("Item3Note"))
            For i = 0 To Item4_1.Items.Count - 1
                If Convert.ToString(dr("Item4_1")) = Item4_1.Items(i).Value Then Item4_1.Items(i).Selected = True
            Next
            Item4Pros.Text = Convert.ToString(dr("Item4Pros"))
            Item4Note.Text = Convert.ToString(dr("Item4Note"))
            For i = 0 To Item5_1.Items.Count - 1
                If Convert.ToString(dr("Item5_1")) = Item5_1.Items(i).Value Then Item5_1.Items(i).Selected = True
            Next
            Item5Pros.Text = Convert.ToString(dr("Item5Pros"))
            Item5Note.Text = Convert.ToString(dr("Item5Note"))
            For i = 0 To Item6_1.Items.Count - 1
                If Convert.ToString(dr("Item6_1")) = Item6_1.Items(i).Value Then Item6_1.Items(i).Selected = True
            Next
            Item6Count1.Text = Convert.ToString(dr("Item6Count1"))
            Item6Count2.Text = Convert.ToString(dr("Item6Count2"))
            Item6Count3.Text = Convert.ToString(dr("Item6Count3"))
            Item6Names.Text = Convert.ToString(dr("Item6Names"))
            Item6Note.Text = Convert.ToString(dr("Item6Note"))
            Item7Note.Text = Convert.ToString(dr("Item7Note"))
            CurseName.Text = Convert.ToString(dr("CurseName"))
            VisitorName.Text = Convert.ToString(dr("VisitorName"))
        End If

        sql = " SELECT b.TMID, b.TrainID, b.TrainName, a.OCID, a.ClassCName, a.CyclType, a.LevelType, a.RID, c.OrgName "
        sql &= " FROM (SELECT * FROM Class_ClassInfo WHERE OCID = '" & OCID & "') a, Key_TrainType b, Auth_Relship d, Org_OrgInfo c "
        sql &= " WHERE a.TMID = b.TMID AND d.OrgID = c.OrgID AND d.RID = a.RID "
        dr = DbAccess.GetOneRow(sql, objconn)

        If Not dr Is Nothing Then
            center.Text = dr("OrgName")
            RIDValue.Value = dr("RID").ToString
            TMID1.Text = "[" & dr("TrainID") & "]" & dr("TrainName")
            TMIDValue1.Value = dr("TMID")
            OCID1.Text = dr("ClassCName")
            If CInt(dr("CyclType")) <> 0 Then OCID1.Text += "第" & CInt(dr("CyclType")) & "期"
            OCIDValue1.Value = dr("OCID")
            center.Enabled = False
            TMID1.Enabled = False
            OCID1.Enabled = False
            Button2.Disabled = True
            Button3.Disabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable
        Dim dr As DataRow

        Dim SeqNo As Integer

        If Request("OCID") = "" Then '表示新增狀態
            If OCIDValue1.Value = "" Then
                Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別")
                Exit Sub
            End If
            '先取出最大SeqNo
            sql = " SELECT MAX(SeqNO) AS num FROM Class_Visitor WHERE OCID = '" & OCIDValue1.Value & "' "
            dr = DbAccess.GetOneRow(sql, objconn)
            If Not dr Is Nothing Then
                If IsDBNull(dr("num")) Then
                    SeqNo = 1
                Else
                    SeqNo = CInt(dr("num")) + 1
                End If
            End If
#Region "(No Use)"

            'sql = "SELECT "
            'sql += " dbo.SUBSTR(ITEM10NOTE, 1, 4000) ITEM10NOTE,ITEM11,dbo.SUBSTR(ITEM11PROS, 1, 4000) ITEM11PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM11NOTE, 1, 4000) ITEM11NOTE,ITEM12,dbo.SUBSTR(ITEM12PROS, 1, 4000) ITEM12PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM12NOTE, 1, 4000) ITEM12NOTE,dbo.SUBSTR(ITEM13NOTE, 1, 4000) ITEM13NOTE,  " & vbCrLf
            'sql += " AHEADJOBCOUNT,STUDYTICKETCOUNT,DATA9,DATACOPY9,DATA9NOTE,DATA10,DATACOPY10,DATA10NOTE,DATA11,DATACOPY11,  " & vbCrLf
            'sql += " DATA11NOTE,DATA12,ITEM14,ITEM14TECH,ITEM15,ITEM16,ITEM17,ITEM18,ITEM19,ITEM20,ITEM21,ITEM22,ITEM23,ITEM24,  " & vbCrLf
            'sql += " ITEM25,ITEM26,ITEM27,ITEM28,ITEM28COUNT,ITEM28_2,ITEM29,dbo.SUBSTR(ITEM29PROS, 1, 4000) ITEM29PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM29NOTE, 1, 4000) ITEM29NOTE,ITEM30,dbo.SUBSTR(ITEM31NOTE, 1, 4000) ITEM31NOTE,  " & vbCrLf
            'sql += " ITEM32,dbo.SUBSTR(ITEM32NOTE, 1, 4000) ITEM32NOTE,D1C,D2C,D3C,D4C,D5C,D6C,D7C,D8C,D3C3,D4C3,D5C3,D6C3,  " & vbCrLf
            'sql += " D7C3,D8C3,OCID,SEQNO,APPLYDATE,AUTHCOUNT,TURTHCOUNT,TURNOUTCOUNT,TRUANCYCOUNT,REJECTCOUNT,DATA1,DATACOPY1,  " & vbCrLf
            'sql += " DATA1NOTE,DATA2,DATACOPY2,DATA2NOTE,DATA3,DATACOPY3,DATA3NOTE,DATA4,DATACOPY4,DATA4NOTE,DATA5,DATACOPY5,  " & vbCrLf
            'sql += " DATA5NOTE,DATA6,DATACOPY6,DATA6NOTE,DATA7,DATACOPY7,DATA7NOTE,ITEM1_1,ITEM1_2,ITEM1_3,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM1PROS, 1, 4000) ITEM1PROS,dbo.SUBSTR(ITEM1NOTE, 1, 4000) ITEM1NOTE,ITEM2_1,  " & vbCrLf
            'sql += " ITEM2_2,dbo.SUBSTR(ITEM2PROS, 1, 4000) ITEM2PROS,dbo.SUBSTR(ITEM2NOTE, 1, 4000) ITEM2NOTE,  " & vbCrLf
            'sql += " ITEM3_1,ITEM3_1TECH,ITEM3_1TUTOR,ITEM3_2,dbo.SUBSTR(ITEM3PROS, 1, 4000) ITEM3PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM3NOTE, 1, 4000) ITEM3NOTE,ITEM4_1,dbo.SUBSTR(ITEM4PROS, 1, 4000) ITEM4PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM4NOTE, 1, 4000) ITEM4NOTE,ITEM5_1,dbo.SUBSTR(ITEM5PROS, 1, 4000) ITEM5PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM5NOTE, 1, 4000) ITEM5NOTE,ITEM6_1,ITEM6COUNT1,ITEM6COUNT2,ITEM6COUNT3,ITEM6NAMES,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM6NOTE, 1, 4000) ITEM6NOTE,dbo.SUBSTR(ITEM7NOTE, 1, 4000) ITEM7NOTE,CURSENAME,  " & vbCrLf
            'sql += " VISITORNAME,MODIFYACCT,MODIFYDATE,RID,ISCLEAR,DATA8,VISITHOUR,TPERIOD,COMCOUNT,BOOKCOUNT,NETCOUNT,PSYCOUNT,  " & vbCrLf
            'sql += " ITEM8,dbo.SUBSTR(ITEM8PROS, 1, 4000) ITEM8PROS,dbo.SUBSTR(ITEM8NOTE, 1, 4000) ITEM8NOTE,  " & vbCrLf
            'sql += " ITEM9,ITEM9_TECH,dbo.SUBSTR(ITEM9PROS, 1, 4000) ITEM9PROS,dbo.SUBSTR(ITEM9NOTE, 1, 4000) ITEM9NOTE,  " & vbCrLf
            'sql += " ITEM10,dbo.SUBSTR(ITEM10PROS, 1, 4000) ITEM10PROS   " & vbCrLf

#End Region
            sql = ""
            sql &= " SELECT * FROM Class_Visitor WHERE 1<>1 " & vbCrLf
            dt = DbAccess.GetDataTable(sql, da, objconn)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = OCIDValue1.Value
            dr("SeqNo") = SeqNo
        Else
#Region "(No Use)"

            'sql = "SELECT "
            'sql += " dbo.SUBSTR(ITEM10NOTE, 1, 4000) ITEM10NOTE,ITEM11,dbo.SUBSTR(ITEM11PROS, 1, 4000) ITEM11PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM11NOTE, 1, 4000) ITEM11NOTE,ITEM12,dbo.SUBSTR(ITEM12PROS, 1, 4000) ITEM12PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM12NOTE, 1, 4000) ITEM12NOTE,dbo.SUBSTR(ITEM13NOTE, 1, 4000) ITEM13NOTE,  " & vbCrLf
            'sql += " AHEADJOBCOUNT,STUDYTICKETCOUNT,DATA9,DATACOPY9,DATA9NOTE,DATA10,DATACOPY10,DATA10NOTE,DATA11,DATACOPY11,  " & vbCrLf
            'sql += " DATA11NOTE,DATA12,ITEM14,ITEM14TECH,ITEM15,ITEM16,ITEM17,ITEM18,ITEM19,ITEM20,ITEM21,ITEM22,ITEM23,ITEM24,  " & vbCrLf
            'sql += " ITEM25,ITEM26,ITEM27,ITEM28,ITEM28COUNT,ITEM28_2,ITEM29,dbo.SUBSTR(ITEM29PROS, 1, 4000) ITEM29PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM29NOTE, 1, 4000) ITEM29NOTE,ITEM30,dbo.SUBSTR(ITEM31NOTE, 1, 4000) ITEM31NOTE,  " & vbCrLf
            'sql += " ITEM32,dbo.SUBSTR(ITEM32NOTE, 1, 4000) ITEM32NOTE,D1C,D2C,D3C,D4C,D5C,D6C,D7C,D8C,D3C3,D4C3,D5C3,D6C3,  " & vbCrLf
            'sql += " D7C3,D8C3,OCID,SEQNO,APPLYDATE,AUTHCOUNT,TURTHCOUNT,TURNOUTCOUNT,TRUANCYCOUNT,REJECTCOUNT,DATA1,DATACOPY1,  " & vbCrLf
            'sql += " DATA1NOTE,DATA2,DATACOPY2,DATA2NOTE,DATA3,DATACOPY3,DATA3NOTE,DATA4,DATACOPY4,DATA4NOTE,DATA5,DATACOPY5,  " & vbCrLf
            'sql += " DATA5NOTE,DATA6,DATACOPY6,DATA6NOTE,DATA7,DATACOPY7,DATA7NOTE,ITEM1_1,ITEM1_2,ITEM1_3,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM1PROS, 1, 4000) ITEM1PROS,dbo.SUBSTR(ITEM1NOTE, 1, 4000) ITEM1NOTE,ITEM2_1,  " & vbCrLf
            'sql += " ITEM2_2,dbo.SUBSTR(ITEM2PROS, 1, 4000) ITEM2PROS,dbo.SUBSTR(ITEM2NOTE, 1, 4000) ITEM2NOTE,  " & vbCrLf
            'sql += " ITEM3_1,ITEM3_1TECH,ITEM3_1TUTOR,ITEM3_2,dbo.SUBSTR(ITEM3PROS, 1, 4000) ITEM3PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM3NOTE, 1, 4000) ITEM3NOTE,ITEM4_1,dbo.SUBSTR(ITEM4PROS, 1, 4000) ITEM4PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM4NOTE, 1, 4000) ITEM4NOTE,ITEM5_1,dbo.SUBSTR(ITEM5PROS, 1, 4000) ITEM5PROS,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM5NOTE, 1, 4000) ITEM5NOTE,ITEM6_1,ITEM6COUNT1,ITEM6COUNT2,ITEM6COUNT3,ITEM6NAMES,  " & vbCrLf
            'sql += " dbo.SUBSTR(ITEM6NOTE, 1, 4000) ITEM6NOTE,dbo.SUBSTR(ITEM7NOTE, 1, 4000) ITEM7NOTE,CURSENAME,  " & vbCrLf
            'sql += " VISITORNAME,MODIFYACCT,MODIFYDATE,RID,ISCLEAR,DATA8,VISITHOUR,TPERIOD,COMCOUNT,BOOKCOUNT,NETCOUNT,PSYCOUNT,  " & vbCrLf
            'sql += " ITEM8,dbo.SUBSTR(ITEM8PROS, 1, 4000) ITEM8PROS,dbo.SUBSTR(ITEM8NOTE, 1, 4000) ITEM8NOTE,  " & vbCrLf
            'sql += " ITEM9,ITEM9_TECH,dbo.SUBSTR(ITEM9PROS, 1, 4000) ITEM9PROS,dbo.SUBSTR(ITEM9NOTE, 1, 4000) ITEM9NOTE,  " & vbCrLf
            'sql += " ITEM10,dbo.SUBSTR(ITEM10PROS, 1, 4000) ITEM10PROS   " & vbCrLf

#End Region
            sql = ""
            sql &= " SELECT * FROM Class_Visitor WHERE OCID = '" & Request("OCID") & "' AND SeqNo = '" & Request("SeqNo") & "' "
            dt = DbAccess.GetDataTable(sql, da, objconn)
            dr = dt.Rows(0)
            OCIDValue1.Value = Request("OCID")
            SeqNo = Request("SeqNo")
        End If

        dr("ApplyDate") = ApplyDate.Text
        If AuthCount.Text <> "" Then dr("AuthCount") = AuthCount.Text
        If TurthCount.Text <> "" Then dr("TurthCount") = TurthCount.Text
        If TurnoutCount.Text <> "" Then dr("TurnoutCount") = TurnoutCount.Text
        If TruancyCount.Text <> "" Then dr("TruancyCount") = TruancyCount.Text
        If RejectCount.Text <> "" Then dr("RejectCount") = RejectCount.Text
        dr("Data1") = Data1.SelectedValue
        If DataCopy1.Text <> "" Then dr("DataCopy1") = DataCopy1.Text
        If Data1Note.Text <> "" Then dr("Data1Note") = Data1Note.Text
        dr("Data2") = Data2.SelectedValue
        If DataCopy2.Text <> "" Then dr("DataCopy2") = DataCopy2.Text
        If Data2Note.Text <> "" Then dr("Data2Note") = Data2Note.Text
        dr("Data3") = Data3.SelectedValue
        If DataCopy3.Text <> "" Then dr("DataCopy3") = DataCopy3.Text
        If Data3Note.Text <> "" Then dr("Data3Note") = Data3Note.Text
        dr("Data4") = Data4.SelectedValue
        If DataCopy4.Text <> "" Then dr("DataCopy4") = DataCopy4.Text
        If Data4Note.Text <> "" Then dr("Data4Note") = Data4Note.Text
        dr("Data5") = Data5.SelectedValue
        If DataCopy5.Text <> "" Then dr("DataCopy5") = DataCopy5.Text
        If Data5Note.Text <> "" Then dr("Data5Note") = Data5Note.Text
        dr("Data6") = Data6.SelectedValue
        If DataCopy6.Text <> "" Then dr("DataCopy6") = DataCopy6.Text
        If Data6Note.Text <> "" Then dr("Data6Note") = Data6Note.Text
        dr("Data7") = Data7.SelectedValue
        If DataCopy7.Text <> "" Then dr("DataCopy7") = DataCopy7.Text
        If Data7Note.Text <> "" Then dr("Data7Note") = Data7Note.Text
        dr("Item1_1") = Item1_1.SelectedValue
        dr("Item1_2") = Item1_2.SelectedValue
        dr("Item1_3") = Item1_3.Text
        If Item1Pros.Text <> "" Then dr("Item1Pros") = Item1Pros.Text
        If Item1Note.Text <> "" Then dr("Item1Note") = Item1Note.Text
        dr("Item2_1") = Item2_1.SelectedValue
        dr("Item2_2") = Item2_2.SelectedValue
        If Item2Pros.Text <> "" Then dr("Item2Pros") = Item2Pros.Text
        If Item2Note.Text <> "" Then dr("Item2Note") = Item2Note.Text
        dr("Item3_1") = Item3_1.SelectedValue
        dr("Item3_1Tech") = Item3_1Tech.Text
        If Item3_1Tutor.Text <> "" Then dr("Item3_1Tutor") = Item3_1Tutor.Text
        dr("Item3_2") = Item3_2.SelectedValue
        If Item3Pros.Text <> "" Then dr("Item3Pros") = Item3Pros.Text
        If Item3Note.Text <> "" Then dr("Item3Note") = Item3Note.Text
        dr("Item4_1") = Item4_1.SelectedValue
        If Item4Pros.Text <> "" Then dr("Item4Pros") = Item4Pros.Text
        If Item4Note.Text <> "" Then dr("Item4Note") = Item4Note.Text
        dr("Item5_1") = Item5_1.SelectedValue
        If Item5Pros.Text <> "" Then dr("Item5Pros") = Item5Pros.Text
        If Item5Note.Text <> "" Then dr("Item5Note") = Item5Note.Text
        dr("Item6_1") = Item6_1.SelectedValue
        If Item6Count1.Text <> "" Then dr("Item6Count1") = Item6Count1.Text
        If Item6Count2.Text <> "" Then dr("Item6Count2") = Item6Count2.Text
        If Item6Count3.Text <> "" Then dr("Item6Count3") = Item6Count3.Text
        If Item6Names.Text <> "" Then dr("Item6Names") = Item6Names.Text
        If Item6Note.Text <> "" Then dr("Item6Note") = Item6Note.Text
        If Item7Note.Text <> "" Then dr("Item7Note") = Item7Note.Text
        dr("CurseName") = CurseName.Text
        dr("VisitorName") = VisitorName.Text
        dr("RID") = sm.UserInfo.RID
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Try
            DbAccess.UpdateDataTable(dt, da)
            If Not Session(cst_SearchStr) Is Nothing Then Session(cst_SearchStr) = Session(cst_SearchStr) 'Else Session(cst_SearchStr) = Me.ViewState("SearchStr")
            If Me.ViewState("_SearchStr") <> "" Then Session(cst_SearchStr2) = Me.ViewState("_SearchStr")
            Common.RespWrite(Me, "<script> alert('儲存成功');")
            If Request("DOCID") <> "" Then
                Common.RespWrite(Me, "location.href='CP_01_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID") & "';</script>")
            Else
                Common.RespWrite(Me, "location.href='CP_01_001.aspx?ID=" & Request("ID") & "';</script>")
            End If
        Catch ex As Exception
            'Common.RespWrite(Me, ex)
            Throw ex
        End Try
    End Sub

    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        If Not Session(cst_SearchStr) Is Nothing Then Session(cst_SearchStr) = Session(cst_SearchStr) 'Else Session(cst_SearchStr) = Me.ViewState("SearchStr")
        If Me.ViewState("_SearchStr") <> "" Then Session(cst_SearchStr2) = Me.ViewState("_SearchStr")
        If Request("DOCID") <> "" Then
            TIMS.Utl_Redirect1(Me, "CP_01_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID"))
        Else
            TIMS.Utl_Redirect1(Me, "CP_01_001.aspx?ID=" & Request("ID"))
        End If
    End Sub

    Private Sub ExtraFunction()
        Dim sql As String
        Dim dr As DataRow
        sql = ""
        sql &= " SELECT COUNT(1) AS num "
        sql &= " FROM class_studentsofclass a "
        sql &= " JOIN Stud_Turnout b ON a.socid = b.socid "
        sql &= " WHERE b.LeaveDate = " & TIMS.To_date(ApplyDate.Text)
        sql &= "    AND a.OCID = '" & OCIDValue1.Value & "' "
        dr = DbAccess.GetOneRow(sql, objconn)

        If dr IsNot Nothing Then
            If IsDBNull(dr("num")) Then
                Data5.SelectedIndex = 0
                Data5Note.Text = Replace(Data5Note.Text, "無人請假", "") & "無人請假"
            ElseIf dr("num") = 0 Then
                Data5.SelectedIndex = 0
                Data5Note.Text = Replace(Data5Note.Text, "無人請假", "") & "無人請假"
            ElseIf dr("num") > 0 Then
                Data5Note.Text = Replace(Data5Note.Text, "無人請假", "")
            End If
        End If

        sql = ""
        sql &= " SELECT COUNT(1) AS num "
        sql &= " FROM class_studentsofclass "
        sql &= " WHERE OCID = '" & OCIDValue1.Value & "' "
        sql &= " AND StudStatus = 3 "
        dr = DbAccess.GetOneRow(sql, objconn)

        If dr IsNot Nothing Then
            If IsDBNull(dr("num")) Then
                Data6.SelectedIndex = 0
                Data6Note.Text = Replace(Data6Note.Text, "無人退訓", "") & "無人退訓"
            ElseIf dr("num") = 0 Then
                Data6.SelectedIndex = 0
                Data6Note.Text = Replace(Data6Note.Text, "無人退訓", "") & "無人退訓"
            ElseIf dr("num") > 0 Then
                Data6Note.Text = Replace(Data6Note.Text, "無人退訓", "")
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If OCIDValue1.Value <> "" Then
            ExtraFunction()
        End If
        FindRange()
    End Sub

    Private Sub FindRange()
        Dim sql As String
        Dim dr As DataRow
        If ApplyDate.Text = "" Then Me.NowDate.Value = Now.Date.ToShortDateString Else Me.NowDate.Value = ApplyDate.Text
        If OCIDValue1.Value = "" Then
            Me.StartDate.Value = "1911/01/01"
            Me.EndDate.Value = "2099/12/31"
        Else
            If OCIDValue1.Value <> "" Then
                sql = ""
                sql &= " SELECT STDate, FTDate FROM class_classinfo WHERE OCID = '" & OCIDValue1.Value & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    If IsDBNull(dr("STDate")) Then Me.StartDate.Value = "1911/01/01" Else Me.StartDate.Value = dr("STDate").ToString
                    If IsDBNull(dr("FTDate")) Then Me.EndDate.Value = "2099/12/31" Else Me.EndDate.Value = dr("FTDate").ToString
                Else
                    Me.StartDate.Value = "1911/01/01"
                    Me.EndDate.Value = "2099/12/31"
                End If
            Else
                Me.StartDate.Value = "1911/01/01"
                Me.EndDate.Value = "2099/12/31"
            End If
        End If
    End Sub
End Class