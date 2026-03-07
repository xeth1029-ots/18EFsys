Partial Class SD_10_006_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "student_inclass"

    Const cst_printFN2 As String = "in_class"

    Const cst_printFN3s As String = "close_21"
    Const cst_printFN3o As String = "close_22"

    'close_21 'close_22
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            Common.SetListItem(Type, "1")
            'Type.SelectedValue = 1
            labMsg.Visible = False

            '在訓'證明字號
            NO1.Text = TIMS.GetGlobalVar(Me, "5", "1", objconn)
            '受訓'證明字號
            NO2.Text = TIMS.GetGlobalVar(Me, "10", "1", objconn)
            '結訓'證明字號
            NO3.Text = TIMS.GetGlobalVar(Me, "11", "1", objconn)
        End If

        '1.在訓證明(student_inclass)/2.受訓證明(in_class)/3.結訓證書
        Dim v_Type As String = TIMS.GetListValue(Type)
        If v_Type = "1" Then
            NO1_TR.Style("display") = ""
            NO2_TR.Style("display") = "none"
            NO3_TR.Style("display") = "none"
            NO3_TR2.Style("display") = "none"
        ElseIf v_Type = "2" Then
            NO1_TR.Style("display") = "none"
            NO2_TR.Style("display") = ""
            NO3_TR.Style("display") = "none"
            NO3_TR2.Style("display") = "none"
        ElseIf v_Type = "3" Then
            NO1_TR.Style("display") = "none"
            NO2_TR.Style("display") = "none"
            NO3_TR.Style("display") = ""
            NO3_TR2.Style("display") = ""
        End If

        Type.Attributes("onclick") = "ShowTR();"
        'Search.Attributes("onclick") = "ChekSearch();"
        Search.Attributes.Add("onclick", "return ChekSearch();")
    End Sub

    Private Sub Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Search.Click
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If IDNO.Text = "" Then
            Common.MessageBox(Me, "請輸入身分證號!!")
            Exit Sub
        End If

        Dim parms As New Hashtable From {
            {"TPLANID", sm.UserInfo.TPlanID},
            {"DISTID", sm.UserInfo.DistID},
            {"IDNO", IDNO.Text}
        }
        Dim sql As String = ""
        sql &= " select cs.studentid" & vbCrLf
        sql &= " ,cc.OCID,cs.SOCID" & vbCrLf
        sql &= " ,ss.Name SName" & vbCrLf
        sql &= " ,ss.IDNO,ss.birthday" & vbCrLf
        sql &= " ,id.Name Distid" & vbCrLf
        sql &= " ,ip.years" & vbCrLf
        sql &= " ,kp.PlanName" & vbCrLf
        sql &= " ,oo.orgName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,cs.STUDSTATUS" & vbCrLf
        sql &= " ,dbo.FN_STUDSTATUS_N(cs.StudStatus) STUDSTATUS_N" & vbCrLf
        sql &= " FROM STUD_STUDENTINFO ss" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on ss.sid = cs.sid" & vbCrLf
        sql &= " join Class_ClassInfo cc on  cc.ocid = cs.ocid" & vbCrLf
        sql &= " join Plan_PlanInfo pp on cc.planid = pp.planid and cc.COMIDNO = pp.COMIDNO and cc.seqno = pp.seqno" & vbCrLf
        sql &= " join ID_Plan ip on pp.planid = ip.planid" & vbCrLf
        sql &= " join ID_District id on ip.distid = id.distid" & vbCrLf
        sql &= " join Org_OrgInfo oo on oo.comidno = pp.comidno" & vbCrLf
        sql &= " join Key_Plan kp on kp.Tplanid = ip.Tplanid" & vbCrLf
        sql &= " where ip.TPLANID=@TPLANID " & vbCrLf
        sql &= " and ip.DISTID =@DISTID" & vbCrLf
        sql &= " and ss.IDNO = @IDNO" & vbCrLf
        sql &= " order by cc.STDate" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            DataGrid1.DataSource = dt
            labMsg.Visible = False
            DataGrid1.DataBind()
        Else
            labMsg.Visible = True
            DataGrid1.Visible = False
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Print As Button = e.Item.FindControl("Print")
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                '0.studentid/1.OCID/2.SOCID/3.STUDSTATUS
                Print.CommandArgument = String.Concat(drv("studentid"), ",", drv("OCID"), ",", drv("SOCID"), ",", drv("STUDSTATUS"))
        End Select

    End Sub

    Function Get_ValidateNG(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs, ByRef arr As Array) As Boolean
        Dim flagErr As Boolean = False

        flagErr = (e.CommandArgument Is Nothing)
        If flagErr Then Return flagErr

        flagErr = (e.CommandArgument = "")
        If flagErr Then Return flagErr

        flagErr = (e.CommandArgument.ToString.IndexOf(",") = -1)
        If flagErr Then Return flagErr

        '0.studentid/1.OCID/2.SOCID/3.STUDSTATUS
        arr = Split(e.CommandArgument, ",")
        flagErr = (arr Is Nothing)
        If flagErr Then Return flagErr

        flagErr = (arr.Length <> 4)
        If flagErr Then Return flagErr

        Return flagErr
    End Function

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim arr As Array = Nothing '0.studentid/1.OCID/2.SOCID/3.STUDSTATUS
        Dim flagErr As Boolean = Get_ValidateNG(e, arr)
        If flagErr Then
            Common.MessageBox(Me, "資料異常，請重新查詢資料！")
            Exit Sub
        End If

        'Dim arr As Array '0.studentid/1.OCID/2.SOCID/3.STUDSTATUS
        '1.在訓證明(student_inclass)/2.受訓證明(in_class)/3.結訓證書
        Dim v_Type As String = TIMS.GetListValue(Type)
        Dim StudentID As String = String.Concat("\'", arr(0), "\'") '0.studentid一次只有一個學員
        Dim STUDSTATUS As String = TIMS.ClearSQM(arr(3)) '3.STUDSTATUS

        Dim v_rblYearType1 As String = TIMS.GetListValue(rblYearType1) '列印格式 1:西元年 2:民國年
        Dim v_RTE As String = If(v_rblYearType1 = "1", "E", "C") '列印格式 E:西元年 C:民國年 #{RTE} RTE

        Dim MyValue As String = ""

        Select Case v_Type'Type.SelectedValue
            Case "1" '1.在訓證明(student_inclass)/2.受訓證明(in_class)/3.結訓證書
                If NO1.Text = "" Then
                    Common.MessageBox(Me, "請輸入在職證明字號。")
                    Exit Sub
                End If
                If STUDSTATUS = "5" Then
                    Common.MessageBox(Me, "該學員已結訓。")
                    Exit Sub
                End If

                'Dim MyValue As String = ""
                MyValue = ""
                MyValue &= String.Concat("&DistID=", sm.UserInfo.DistID)
                MyValue &= String.Concat("&OCID=", Val(arr(1)))
                MyValue &= String.Concat("&StudentID=", StudentID)
                MyValue &= String.Concat("&ProveNum=", NO1.Text)
                MyValue &= String.Concat("&rblYearType1=", v_rblYearType1) '列印格式 1:西元年 2:民國年
                MyValue &= String.Concat("&RTE=", v_RTE) '列印格式 E:西元年 C:民國年 #{RTE} RTE
                MyValue &= "&Type=1" '$P{Type}=="1"?$F{PN}+" 補發":$F{PN}
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

            Case "2" '1.在訓證明(student_inclass)/2.受訓證明(in_class)/3.結訓證書
                If NO2.Text = "" Then
                    Common.MessageBox(Me, "請輸入受訓證明字號。")
                    Exit Sub
                End If
                If STUDSTATUS <> "5" Then
                    Common.MessageBox(Me, "該學員未結訓。")
                    Exit Sub
                End If

                'Dim v_rblYearType1 As String = TIMS.GetListValue(rblYearType1) '列印格式 1:西元年 2:民國年
                'Dim v_RTE As String = If(v_rblYearType1 = "1", "E", "C") '列印格式 E:西元年 C:民國年 #{RTE} RTE
                'Dim MyValue As String = ""
                MyValue = ""
                MyValue &= String.Concat("&DistID=", sm.UserInfo.DistID)
                MyValue &= String.Concat("&OCID=", Val(arr(1)))
                MyValue &= String.Concat("&StudentID=", StudentID)
                MyValue &= String.Concat("&ProveNum=", NO2.Text)
                MyValue &= String.Concat("&rblYearType1=", v_rblYearType1) '列印格式 1:西元年 2:民國年
                MyValue &= String.Concat("&RTE=", v_RTE) '列印格式 E:西元年 C:民國年 #{RTE} RTE
                MyValue &= "&Type=1" '$P{Type}=="1"?$F{PN}+" 補發":$F{PN}
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)

            Case "3" '3.結訓證書
                Dim classCnt As Integer = 0
                '該班 訓練課程與授課時數資料(實際排課授課時數)
                If Convert.ToString(arr(1)) <> "" Then
                    'MVIEW_CLASS_SCHEDULE
                    Dim pms1 As New Hashtable From {{"OCID", Val(arr(1))}}
                    Dim sql As String = ""
                    sql &= " SELECT DISTINCT SC.OCID,SC.COURSEID,SC.COURSENAME" & vbCrLf
                    sql &= " FROM MVIEW_CLASS_SCHEDULE sc " & vbCrLf
                    sql &= " WHERE sc.OCID=@OCID" & vbCrLf
                    Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)
                    If TIMS.dtHaveDATA(dt) Then classCnt = dt.Rows.Count
                End If
                If classCnt = 0 Then
                    Common.MessageBox(Me, "該班無訓練課程與授課時數資料!")
                    Exit Sub
                End If
                If NO3.Text = "" Then
                    Common.MessageBox(Me, "請輸入結訓證明字號。")
                    Exit Sub
                End If

                MyValue = ""
                MyValue &= "&DistID=" & sm.UserInfo.DistID
                MyValue &= "&StudentID=" & StudentID
                MyValue &= "&OCID=" & Val(arr(1))
                MyValue &= "&ProveNum=" & Convert.ToString(NO3.Text)
                MyValue &= "&rblYearType1=" & v_rblYearType1 '列印格式 1:西元年 2:民國年
                MyValue &= "&Type=1" '$P{Type}=="1"?$F{PN}+" 補發":$F{PN}

                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "close", MyValue)\
                'cst_printFN3s:自辦 // 'cst_printFN3o:委訓
                Dim sPrintFN As String = If(Type2.SelectedValue = "1", cst_printFN3s, cst_printFN3o)

                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrintFN, MyValue)

        End Select
    End Sub
End Class
