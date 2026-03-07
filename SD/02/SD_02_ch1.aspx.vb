Partial Class SD_02_ch1
    Inherits AuthBasePage

    'Dim sql As String
    'Dim MySQLStr As String
    'Dim dt As DataTable
    'Dim i As Integer
    'Dim Key_TrainType As DataTable
    'Dim PageControler1 As New PageControler

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        TIMS.ChkSession(Me, 9, sm)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        'PageControler1 = Me.FindControl("PageControler1")
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'Request("AcctPlan")=1              表示可以跨計畫選擇
        'Request("special")=1               增加回傳開訓跟結訓日期
        'Request("special")=2               母頁Submit
        'Request("special")=3               SD_05_005.aspx專用
        'Request("sort")=3                  表示回傳到TMID3.value
        'msg.Text = ""
        search_but.Attributes("onclick") = "javascript:return chkdata();"

        'Sql = "select * from Key_TrainType"
        'Key_TrainType = DbAccess.GetDataTable(Sql)

        If Not IsPostBack Then
            msg.Text = ""
            Table2.Visible = False

            syear = TIMS.GetSyear(syear)
            Common.SetListItem(syear, Now.Year)
            hourran = TIMS.GET_HOURRAN(hourran, objconn, sm)

            If Request.Cookies("SD_syear") IsNot Nothing Then
                Common.SetListItem(syear, Request.Cookies("SD_syear").Value)
                tb_career_id.Text = Request.Cookies("SD_TB_career_id").Value
                trainvalue.Value = Request.Cookies("SD_trainValue").Value
                Common.SetListItem(hourran, Request.Cookies("SD_HourRan").Value)
                cycltype.Text = Request.Cookies("SD_CyclType").Value
                classid.Text = Request.Cookies("SD_ClassID").Value
                Common.SetListItem(classround, Request.Cookies("SD_ClassRound").Value)
                'search_but_Click(sender, e)
                Call Search1()
            End If
        End If
    End Sub

    Sub GetSearch()
        Response.Cookies("SD_syear").Value = TIMS.ClearSQM(syear.SelectedValue)
        Response.Cookies("SD_TB_career_id").Value = TIMS.ClearSQM(tb_career_id.Text)
        Response.Cookies("SD_trainValue").Value = TIMS.ClearSQM(trainvalue.Value)
        Response.Cookies("SD_HourRan").Value = TIMS.ClearSQM(hourran.SelectedValue)
        Response.Cookies("SD_CyclType").Value = TIMS.ClearSQM(cycltype.Text)
        Response.Cookies("SD_ClassID").Value = TIMS.ClearSQM(classid.Text)
        Response.Cookies("SD_ClassRound").Value = TIMS.ClearSQM(classround.SelectedValue)
    End Sub

    Sub Search1()
        'Dim TMIDStr As String = ""
        'Dim SearchStr As String = ""
        Call GetSearch()

        'Select Case classround.SelectedIndex
        '    Case 0 '開訓二週前
        '        SearchStr += " and dbo.TRUNC_DATETIME(STDate)-dbo.TRUNC_DATETIME(getdate())<=(2*7) and dbo.TRUNC_DATETIME(STDate)-dbo.TRUNC_DATETIME(getdate())>0" & vbCrLf
        '    Case 1 '已開訓('己開訓改成除排己經結訓的班級)
        '        SearchStr += " and dbo.TRUNC_DATETIME(STDate)-dbo.TRUNC_DATETIME(getdate())<=0 and dbo.TRUNC_DATETIME(FTDate)-dbo.TRUNC_DATETIME(getdate())>0" & vbCrLf
        '    Case 2 '已結訓
        '        SearchStr += " and dbo.TRUNC_DATETIME(FTDate)-dbo.TRUNC_DATETIME(getdate())<=0 " & vbCrLf
        '    Case 3 '未開訓
        '        SearchStr += " and dbo.TRUNC_DATETIME(STDate)-dbo.TRUNC_DATETIME(getdate())>0" & vbCrLf
        '    Case 4 '全部
        'End Select

        'Dim MySQLStr As String = ""
        Dim sql As String = ""
        sql &= " SELECT cc.Years,cc.PlanID,cc.OCID,cc.ClassCName" & vbCrLf
        sql &= " ,format(cc.STDate,'yyyy/MM/dd') STDate" & vbCrLf
        sql &= " ,format(cc.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf
        'sql &= " ,cc.STDate,cc.FTDate" & vbCrLf
        sql &= " ,cc.TMID,cc.IsApplic,cc.CLSID,cc.CyclType" & vbCrLf
        sql &= " ,cc.ComIDNO,cc.SeqNO,cc.LevelType" & vbCrLf
        sql &= " ,cc.TPeriod" & vbCrLf
        sql &= " ,ip.TPlanID" & vbCrLf
        'CYCLTYPE 提供期別可以不填寫，但未填補01 維持一致性
        sql &= " ,concat(cc.YEARS,'0',ISNULL(ic.ClassID2,ic.ClassID),ISNULL(cc.CYCLTYPE,'01')) ClassID " & vbCrLf
        sql &= " ,t.TrainID,t.TrainName" & vbCrLf
        sql &= " ,'['+IsNull(t.TrainID ,t.JobID)+ ']'+IsNull(t.TrainName ,t.JobName) TrainName2 " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,pp.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.PlanID = cc.PlanID AND pp.COMIDNO = cc.COMIDNO AND pp.SEQNO = cc.SEQNO " & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.planid =cc.planid" & vbCrLf
        sql &= " JOIN ID_CLASS c on c.CLSID=cc.CLSID" & vbCrLf
        If sm.UserInfo.RID <> "A" Then '非署限定班級資訊
            sql &= " JOIN (select OCID FROM AUTH_ACCRWCLASS WHERE lower(Account)=lower('" & sm.UserInfo.UserID & "')) auc on auc.OCID=cc.OCID" & vbCrLf
        End If
        sql &= " LEFT JOIN KEY_TRAINTYPE t on t.TMID=cc.TMID" & vbCrLf
        sql &= " WHERE cc.IsSuccess='Y'" & vbCrLf
        sql &= " and cc.NotOpen='N'" & vbCrLf
        sql &= " and ip.TPlanID ='" & sm.UserInfo.TPlanID & "'" '限定登入計畫
        If sm.UserInfo.RID <> "A" Then
            Dim R_AcctPlan As String = TIMS.ClearSQM(Request("AcctPlan")) '1:接受全部計畫
            If R_AcctPlan <> "1" Then
                sql &= " and cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf '限定登入計畫
            End If
        End If

        '班級範圍
        Select Case classround.SelectedIndex
            Case 0 '開訓二週前
                sql &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) <= (2*7) AND DATEDIFF(DAY,GETDATE(),cc.STDate) >=0" & vbCrLf
            Case 1 '已開訓('己開訓改成除排己經結訓的班級)
                sql &= " AND DATEDIFF(DAY,cc.STDate,GETDATE()) >=0 AND DATEDIFF(DAY,cc.FTDate,GETDATE()) < 0" & vbCrLf
            Case 2 '已結訓
                sql &= " AND DATEDIFF(DAY,cc.FTDate,GETDATE()) >=0" & vbCrLf
            Case 3 '未開訓
                sql &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) > 0 " & vbCrLf
            Case 4 '全部
            Case Else '異常
                sql &= " AND 1<>1 " & vbCrLf
        End Select
        sql &= " and cc.Years='" & Right(syear.SelectedValue, 2) & "'" & vbCrLf

        trainvalue.Value = TIMS.ClearSQM(trainvalue.Value)
        If trainvalue.Value <> "" Then
            sql &= " and t.TMID='" & trainvalue.Value & "'" & vbCrLf
        End If
        If hourran.SelectedIndex <> 0 Then
            sql += " and cc.TPeriod='" & hourran.SelectedValue & "'" & vbCrLf
        End If
        Dim str_RID As String = TIMS.ClearSQM(Request("RID"))
        If str_RID = "" Then str_RID = sm.UserInfo.RID
        sql += " and cc.RID='" & str_RID & "'" & vbCrLf

        '期別
        cycltype.Text = TIMS.FmtCyclType(cycltype.Text)
        If cycltype.Text <> "" Then
            If IsNumeric(cycltype.Text) Then
                If Int(cycltype.Text) < 10 Then
                    sql &= " AND cc.CyclType = '0" & Int(cycltype.Text) & "' " & vbCrLf
                Else
                    sql &= " AND cc.CyclType = '" & cycltype.Text & "' " & vbCrLf
                End If
            End If
        End If

        'CYCLTYPE 提供期別可以不填寫，但未填補01 維持一致性
        classid.Text = TIMS.ClearSQM(classid.Text)
        If classid.Text <> "" Then
            sql += " and concat(cc.YEARS,'0',ISNULL(ic.ClassID2,ic.ClassID),ISNULL(cc.CYCLTYPE,'01')) Like '%" & classid.Text & "%'" & vbCrLf
        End If
        classcname.Text = TIMS.ClearSQM(classcname.Text)
        If classcname.Text <> "" Then
            sql += " and cc.ClassCName like '%" + classcname.Text + "%'" & vbCrLf
        End If

        msg.Text = "查無資料!!"
        Table2.Visible = False

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Return

        'If dt.Rows.Count > 0 Then End If
        msg.Text = ""
        Table2.Visible = True

        'PageControler1.SqlString = MySQLStr
        If Hid_SSSDTRID.Value = "" Then Hid_SSSDTRID.Value = TIMS.GetRnd6Eng()
        PageControler1.SSSDTRID = Hid_SSSDTRID.Value 'TIMS.GetRnd6Eng()
        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "OCID"
        PageControler1.ControlerLoad()
    End Sub

    Private Sub Search_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search_but.Click
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim IsApplic_N As String = If(Convert.ToString(drv("IsApplic")) = "Y", "可挑選志願", "不可挑選志願")
                e.Item.Cells(5).Text = IsApplic_N '"可挑選志願"  "不可挑選志願"
        End Select
    End Sub

    Private Sub Send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click
        Dim rqclass1 As String = TIMS.ClearSQM(Request("class1"))
        If rqclass1 = "" Then
            Common.MessageBox(Me, "請先勾選班級!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(rqclass1, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "查無資料!")
            Exit Sub
        End If

        Common.RespWrite(Me, "<script language=javascript>")
        Common.RespWrite(Me, "function returnNum(){")
        If Request("sort") = "" Then
            Common.RespWrite(Me, "window.opener.document.form1.TMID1.value='[" & drCC("TrainID") & "]" & drCC("TrainName") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.TMIDValue1.value='" & drCC("TMID") & "';")
        Else
            Common.RespWrite(Me, "window.opener.document.form1.TMID" & Request("sort") & ".value='[" & drCC("TrainID") & "]" & drCC("TrainName") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.TMIDValue" & Request("sort") & ".value='" & drCC("TMID") & "';")
        End If

        Dim s_className2 As String = Convert.ToString(drCC("CLASSCNAME2"))

        If Request("sort") = "" Then
            Common.RespWrite(Me, "window.opener.document.form1.OCID1.value='" & s_className2 & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCIDValue1.value='" & drCC("OCID") & "';")
        Else
            Common.RespWrite(Me, "window.opener.document.form1.OCID" & Request("sort") & ".value='" & s_className2 & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCIDValue" & Request("sort") & ".value='" & drCC("OCID") & "';")
        End If

        Select Case Request("special")
            Case 1                                  '增加回傳開訓與結訓日期
                Common.RespWrite(Me, "window.opener.document.form1.SDate.value='" & drCC("STDate") & "';")
                Common.RespWrite(Me, "window.opener.document.form1.FDate.value='" & drCC("FTDate") & "';")
            Case 2
                Common.RespWrite(Me, "window.opener.document.form1.submit();")
            Case 3 'SD_05_005.aspx專用
                'Dim FunDr As DataRow
                'Dim FunDt As DataTable = sm.UserInfo.FunDt
                'Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
                'If FunDrArray.Length <> 0 Then
                '    FunDr = FunDrArray(0)
                '    If FunDr("Sech") = 1 Then
                '        'Common.RespWrite(Me, "window.opener.document.form1.Button5.disabled=false;")
                '    End If

                '    If FunDr("Adds") = 1 Then
                '        Common.RespWrite(Me, "window.opener.document.form1.Button4.disabled=false;")
                '        Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=false;")
                '        Common.RespWrite(Me, "window.opener.document.form1.Usered.value='1';")
                '    End If
                'End If
                Common.RespWrite(Me, "window.opener.document.form1.Button4.disabled=false;")
                Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=false;")
                Common.RespWrite(Me, "window.opener.document.form1.Usered.value='1';")
        End Select

        Common.RespWrite(Me, "window.close();")
        Common.RespWrite(Me, "}")
        Common.RespWrite(Me, "returnNum();")
        Common.RespWrite(Me, "</script>")

        For i As Integer = 1 To 10
            If Request.Cookies("Temp_OCID" & i) Is Nothing Then
                Response.Cookies("Temp_OCID" & i).Value = drCC("OCID")
                Response.Cookies("Temp_ClassName" & i).Value = s_className2
                Response.Cookies("Temp_TMID" & i).Value = drCC("TMID")
                Response.Cookies("Temp_TrainName" & i).Value = "[" & drCC("TrainID") & "]" & drCC("TrainName")
                Exit For
            Else
                If i = 10 Then
                    For j As Integer = 1 To 9
                        Response.Cookies("Temp_OCID" & j).Value = Response.Cookies("Temp_OCID" & j + 1).Value
                        Response.Cookies("Temp_ClassName" & j + 1).Value = Response.Cookies("Temp_ClassName" & j + 1).Value
                        Response.Cookies("Temp_TMID" & j).Value = Response.Cookies("Temp_TMID" & j + 1).Value
                        Response.Cookies("Temp_TrainName" & j).Value = Response.Cookies("Temp_TrainName" & j + 1).Value
                    Next

                    Response.Cookies("Temp_OCID" & i).Value = drCC("OCID")
                    Response.Cookies("Temp_ClassName" & i).Value = s_className2
                    Response.Cookies("Temp_TMID" & i).Value = drCC("TMID")
                    Response.Cookies("Temp_TrainName" & i).Value = "[" & drCC("TrainID") & "]" & drCC("TrainName")
                End If
            End If
        Next

    End Sub
End Class
