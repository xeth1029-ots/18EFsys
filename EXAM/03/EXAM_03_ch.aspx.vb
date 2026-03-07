Partial Class EXAM_03_ch
    Inherits AuthBasePage

    Dim sql As String
    Dim Key_TrainType As DataTable
    Dim PlanKind As Integer
    Dim PageControler1 As New PageControler

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9) '☆
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        PageControler1 = Me.FindControl("PageControler1")

        '檢查Session是否存在 End
        'Request("AcctPlan")=1              表示可以跨計畫選擇
        'Request("special")=1               增加回傳開訓跟結訓日期
        'Request("special")=2               母頁Submit
        'Request("special")=3               SD_05_005.aspx專用
        'Request("sort")=3                  表示回傳到TMID3.value
        'request("RWClass")=1               被班級計畫賦予限制
        PlanKind = TIMS.Get_PlanKind(Me, objconn)
        msg.Text = ""
        search_but.Attributes("onclick") = "javascript:return chkdata();"

        If Not IsPostBack Then
            Years = TIMS.GetSyear(Years)
            Common.SetListItem(Years, sm.UserInfo.Years)
            HourRan = TIMS.GET_HOURRAN(HourRan, objconn, sm)
            Table2.Visible = False
        End If

        sql = "select * from Key_TrainType"
        Key_TrainType = DbAccess.GetDataTable(sql, objconn)

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            Dim dt As DataTable = TIMS.GetCookieTable(Me, objconn)
            If dt.Select("ItemName='SD_TB_career_id'").Length <> 0 Then
                TB_career_id.Text = dt.Select("ItemName='SD_TB_career_id'")(0)("ItemValue")
                trainValue.Value = dt.Select("ItemName='SD_trainValue'")(0)("ItemValue")
                Common.SetListItem(HourRan, dt.Select("ItemName='SD_HourRan'")(0)("ItemValue"))
                CyclType.Text = dt.Select("ItemName='SD_CyclType'")(0)("ItemValue")
                ClassID.Text = dt.Select("ItemName='SD_ClassID'")(0)("ItemValue")
                Common.SetListItem(ClassRound, dt.Select("ItemName='SD_ClassRound'")(0)("ItemValue"))
                search_but_Click(sender, e)
            End If
        End If
    End Sub

    Sub GetSearch()
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        TIMS.InsertCookieTable(Me, dt, da, "SD_TB_career_id", TB_career_id.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_trainValue", trainValue.Value, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_HourRan", HourRan.SelectedValue, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_CyclType", CyclType.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassID", ClassID.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassRound", ClassRound.SelectedValue, True, objconn)
    End Sub

    Private Sub search_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search_but.Click
        Dim MySQLStr As String = ""
        Dim TMIDStr As String = ""
        Dim SearchStr As String = ""
        Call GetSearch()

        If Years.SelectedIndex <> 0 Then
            SearchStr += " and Years='" & Right(Years.SelectedValue, 2) & "'" & vbCrLf
        End If

        '班級範圍
        Select Case ClassRound.SelectedIndex
            Case 0 '開訓二週前
                SearchStr &= " AND DATEDIFF(DAY,GETDATE(),STDate) <= (2*7) AND DATEDIFF(DAY,GETDATE(),STDate) >0" & vbCrLf
            Case 1 '已開訓('己開訓改成除排己經結訓的班級)
                SearchStr &= " AND DATEDIFF(DAY,STDate,GETDATE()) >=0 AND DATEDIFF(DAY,FTDate,GETDATE()) < 0" & vbCrLf
            Case 2 '已結訓
                SearchStr &= " AND DATEDIFF(DAY,FTDate,GETDATE()) >=0" & vbCrLf
            Case 3 '未開訓
                SearchStr &= " AND DATEDIFF(DAY,GETDATE(),STDate) > 0 " & vbCrLf
            Case 4 '全部
            Case Else '異常
                SearchStr &= " AND 1<>1 "
        End Select

        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        Dim V_HourRan As String = TIMS.ClearSQM(HourRan.SelectedValue)
        'Dim rq_RID As String = TIMS.ClearSQM(Request("RID"))
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        ClassID.Text = TIMS.ClearSQM(ClassID.Text)

        If trainValue.Value <> "" Then
            SearchStr += " and TMID='" & trainValue.Value & "'" & vbCrLf
            TMIDStr = " and TMID='" & trainValue.Value & "'" & vbCrLf
        End If
        If HourRan.SelectedIndex <> 0 AndAlso V_HourRan <> "" Then
            SearchStr += " and TPeriod='" & V_HourRan & "'" & vbCrLf
        End If

        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                If Int(CyclType.Text) < 10 Then
                    SearchStr += " and CyclType='0" & Int(CyclType.Text) & "'" & vbCrLf
                Else
                    SearchStr += " and CyclType='" & CyclType.Text & "'" & vbCrLf
                End If
            End If
        End If

        If ClassID.Text <> "" Then
            If ClassID.Text.Length = 9 Then
                ClassID.Text = Mid(ClassID.Text, 4, 4)
            End If
            SearchStr += " and CLSID IN (SELECT CLSID FROM ID_Class WHERE ClassID Like'%" & ClassID.Text & "%')" & vbCrLf
        End If

        If sm.UserInfo.RID = "A" Then
            MySQLStr = "" & vbCrLf
            MySQLStr += " select a.*" & vbCrLf
            MySQLStr += " , CASE when t.JobID is null then t.TrainID else t.JobID end   TrainID" & vbCrLf
            MySQLStr += " , CASE when t.JobID is null then t.trainName else t.JobName end TrainName " & vbCrLf
            MySQLStr += " from " & vbCrLf
            MySQLStr += "(select i.*,c.TPlanID,i.Years+'0'+c.ClassID+i.CyclType as ClassID from " & vbCrLf
            MySQLStr += "(select Years,PlanID,OCID,ClassCName,STDate,FTDate,TMID,IsApplic,CLSID,CyclType,ComIDNO,SeqNO,LevelType FROM Class_ClassInfo where IsSuccess='Y' and NotOpen='N'" & SearchStr & " ) i " & vbCrLf
            MySQLStr += "Join(select CLSID as CLSID1,TPlanID,ClassID from ID_Class) c on i.CLSID=c.CLSID1 " & vbCrLf
            MySQLStr += ") a " & vbCrLf
            MySQLStr += "left join (select TMID,TrainID,TrainName,JobID,JobName from Key_TrainType where 1=1" & TMIDStr & ") t on t.TMID=a.TMID" & vbCrLf
        Else
            MySQLStr = "" & vbCrLf
            MySQLStr += " select a.*" & vbCrLf
            MySQLStr += " , CASE when t.JobID is null then t.TrainID else t.JobID end   TrainID" & vbCrLf
            MySQLStr += " , CASE when t.JobID is null then t.trainName else t.JobName end TrainName " & vbCrLf
            MySQLStr += " from " & vbCrLf
            MySQLStr += "(select i.*,c.TPlanID,i.Years+'0'+c.ClassID+i.CyclType as ClassID from " & vbCrLf
            If Request("AcctPlan") = "1" Then
                MySQLStr += "(select Years,PlanID,OCID,ClassCName,STDate,FTDate,TMID,IsApplic,CLSID,CyclType,ComIDNO,SeqNO,LevelType FROM Class_ClassInfo where IsSuccess='Y' and NotOpen='N'" & SearchStr & " ) i " & vbCrLf
            Else
                If sm.UserInfo.DistID = "002" And sm.UserInfo.TPlanID = "36" Then '星光幫控制
                    MySQLStr += "(select Years,PlanID,OCID,ClassCName,STDate,FTDate,TMID,IsApplic,CLSID,CyclType,ComIDNO,SeqNO,LevelType FROM Class_ClassInfo where IsSuccess='Y' and NotOpen='N'" & SearchStr & ") i " & vbCrLf
                Else
                    MySQLStr += "(select Years,PlanID,OCID,ClassCName,STDate,FTDate,TMID,IsApplic,CLSID,CyclType,ComIDNO,SeqNO,LevelType FROM Class_ClassInfo where IsSuccess='Y' and NotOpen='N'" & SearchStr & " ) i " & vbCrLf
                End If
            End If

            MySQLStr += "Join (select CLSID as CLSID1,TPlanID,ClassID from ID_Class where DistID = '" & sm.UserInfo.DistID & "')  c on i.CLSID=c.CLSID1 " & vbCrLf

            MySQLStr += ") a "
            MySQLStr += "left join (select TMID,TrainID,TrainName,JobID,JobName from Key_TrainType where 1=1" & TMIDStr & ") t on t.TMID=a.TMID " & vbCrLf


            If PlanKind = 1 And Request("RWClass") = 1 Then
                MySQLStr += "join (select OCID from Auth_AccRWClass where Account='" & sm.UserInfo.UserID & "') c on a.OCID=c.OCID" & vbCrLf
            End If
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(MySQLStr, objconn)

        msg.Text = "查無資料!!"
        Table2.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table2.Visible = True

            'PageControler1.SqlString = MySQLStr
            PageControler1.PageDataTable = dt '.SqlString = MySQLStr
            PageControler1.PrimaryKey = "OCID"
            PageControler1.Sort = "ClassID,CyclType"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            e.Item.Cells(1).Text = "[" & drv("TrainID").ToString & "]" & drv("TrainName").ToString

            e.Item.Cells(3).Text = TIMS.GET_CLASSNAME(Convert.ToString(drv("ClassCName")), Convert.ToString(drv("CyclType")))

            If e.Item.Cells(5).Text = "Y" Or e.Item.Cells(5).Text = "y" Then
                e.Item.Cells(5).Text = "可挑選志願"
            Else
                e.Item.Cells(5).Text = "不可挑選志願"
            End If
        End If
    End Sub

    Private Sub send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click
        Dim className As String = ""
        Dim sql As String = ""
        Dim dr As DataRow

        If Request("class1") <> "" Then
            sql = "" & vbCrLf
            sql += " SELECT a.* " & vbCrLf
            sql += "  , CASE when b.JobID is null then b.TrainID else b.JobID end   TrainID" & vbCrLf
            sql += "  , CASE when b.JobID is null then b.trainName else b.JobName end TrainName " & vbCrLf
            sql += "  , e.OrgName " & vbCrLf
            sql += " FROM Class_ClassInfo a " & vbCrLf
            sql += " JOIN Key_TrainType b on a.TMID=b.TMID " & vbCrLf
            sql += " JOIN Auth_Relship d on a.RID=d.RID " & vbCrLf
            sql += " JOIN Org_OrgInfo e on e.OrgID=d.OrgID " & vbCrLf
            sql += " WHERE OCID='" & Request("class1") & "'" & vbCrLf
            dr = DbAccess.GetOneRow(sql, objconn)

            If dr Is Nothing Then
                Common.MessageBox(Me, "查無資料!")
            Else
                Common.RespWrite(Me, "<script language=javascript>")
                Common.RespWrite(Me, "function returnNum(){")

                className = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))

                If dr("examdate").ToString <> "" Then
                    Me.ViewState("examdate") = Common.FormatDate(dr("examdate"))
                Else
                    Me.ViewState("examdate") = "未填寫"
                End If

                Select Case Request("sort")
                    Case "1"
                        Common.RespWrite(Me, "window.opener.document.form1.OCID" & Request("sort") & ".value='" & className & "';")
                        Common.RespWrite(Me, "window.opener.document.form1.OCIDValue" & Request("sort") & ".value='" & dr("OCID") & "';")
                        Common.RespWrite(Me, "window.opener.document.form1.TMID" & Request("sort") & ".value='[" & dr("TrainID") & "]" & dr("TrainName") & "';")
                        Common.RespWrite(Me, "window.opener.document.form1.TMIDValue" & Request("sort") & ".value='" & dr("TMID") & "';")
                    Case "2"
                        Common.RespWrite(Me, "window.opener.document.form1.OCID" & Request("sort") & ".value='" & className & "';")
                        Common.RespWrite(Me, "window.opener.document.form1.OCIDValue" & Request("sort") & ".value='" & dr("OCID") & "';")
                        Common.RespWrite(Me, "window.opener.document.form1.TMID" & Request("sort") & ".value='[" & dr("TrainID") & "]" & dr("TrainName") & "';")
                        Common.RespWrite(Me, "window.opener.document.form1.TMIDValue" & Request("sort") & ".value='" & dr("TMID") & "';")
                        Common.RespWrite(Me, "window.opener.document.form1.txt_examdate1.value='" & Me.ViewState("examdate") & "';")
                End Select
                Common.RespWrite(Me, "window.close();")
                Common.RespWrite(Me, "}")
                Common.RespWrite(Me, "returnNum();")
                Common.RespWrite(Me, "</script>")
            End If
        Else
            Common.MessageBox(Me, "請先勾選班級!")
        End If
    End Sub

End Class
