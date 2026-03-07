Public Class SD_05_033
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_05_033_R"

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGridC1

        If Not IsPostBack Then
            tbSearch1.Visible = True
            DataGridC1.Visible = False
            PageControler1.Visible = False
            DataGridS1.Visible = False
            'btnSave1.Visible = False
            labmsg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            btnSave1.Visible = False
            btnBack1.Visible = False
        End If

        'btnSave1.Enabled = True '新增
        'If Not blnCanAdds Then
        '    btnSave1.Enabled = False '新增
        '    TIMS.Tooltip(btnSave1, "(Adds)無權限使用該功能", True)
        'End If
        'btnSearch1.Enabled = True '查詢
        'If Not blnCanAdds Then
        '    btnSearch1.Enabled = False '查詢
        '    TIMS.Tooltip(btnSave1, "(Sech)無權限使用該功能", True)
        'End If

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
        '       'Dim FunDr As DataRow = FunDrArray(0)
        '        If FunDr("Adds") = 1 Then
        '            btnSave1.Enabled = True '新增
        '        Else
        '            btnSave1.Enabled = False
        '            TIMS.Tooltip(btnSave1, "(Adds)無權限使用該功能", True)
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            btnSearch1.Enabled = True '查詢
        '        Else
        '            btnSearch1.Enabled = False
        '            TIMS.Tooltip(btnSearch1, "(Sech) 無權限使用該功能", True)
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        btnSetLevOrg.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        btnSearch1.Attributes("onclick") = "javascript:return search1();"
        'btnAdd1.Attributes("onclick") = "javascript:return search1();"
        'tExamKind.Attributes("onblur") = "Get_Exam(this,'" & tExamName.ClientID & "');"
    End Sub

    '班級查詢sql
    Sub Search1()
        Dim sql As String = ""
        sql &= " select cc.planid ,cc.comidno,cc.seqno" & vbCrLf
        sql &= " ,cc.OCID,cc.RID,oo.ORGNAME" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        sql &= " ,ip.years" & vbCrLf
        sql &= " ,ip.distid" & vbCrLf
        sql &= " ,ip.tplanid" & vbCrLf
        sql &= " FROM Class_ClassInfo cc" & vbCrLf
        sql &= " JOIN Plan_PlanInfo pp ON cc.PlanID = pp.PlanID AND cc.ComIDNO = pp.ComIDNO AND cc.SeqNO = pp.SeqNO" & vbCrLf
        sql &= " JOIN id_class ic on ic.clsid =cc.clsid" & vbCrLf
        sql &= " join id_plan ip on ip.planid =cc.planid" & vbCrLf
        sql &= " join Key_TrainType tt on tt.TMID=cc.TMID" & vbCrLf
        sql &= " join Auth_Relship rr on rr.RID=cc.RID" & vbCrLf
        sql &= " join Org_OrgInfo oo on oo.comidno=cc.comidno" & vbCrLf
        sql &= " WHERE cc.IsSuccess = 'Y'" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        sql &= " and ip.tplanid ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " and ip.years='" & sm.UserInfo.Years & "'" & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql &= " and ip.distid ='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then
            sql &= " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
        Else
            sql &= " and cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        End If
        If OCIDValue1.Value <> "" Then
            sql &= " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)
        labmsg.Text = "查無資料"
        DataGridC1.Visible = False
        If dt.Rows.Count > 0 Then
            'CPdt = dt.Copy()
            labmsg.Text = ""
            DataGridC1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call Search1()
    End Sub

    '學員list sql
    Sub show_student(ByVal vOCID As String)
        Hidocid.Value = vOCID

        tbSearch1.Visible = False
        DataGridC1.Visible = False
        PageControler1.Visible = False

        Dim sql As String = ""
        sql &= " select cs.socid" & vbCrLf
        sql &= " ,cc.planid" & vbCrLf
        sql &= " ,cc.comidno" & vbCrLf
        sql &= " ,cc.seqno" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,oo.orgname" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        sql &= " ,ip.years" & vbCrLf
        sql &= " ,ip.distid" & vbCrLf
        sql &= " ,ip.tplanid" & vbCrLf
        sql += " ,substr(cs.StudentID,-2) StudID" & vbCrLf
        sql += " ,ss.name" & vbCrLf
        sql += " ,ss.idno" & vbCrLf
        sql += " ,CONVERT(varchar, ss.birthday, 111) birthday " & vbCrLf
        sql += " ,dbo.DECODE6(ss.Sex,'M','男','F','女',ss.Sex) Sex2 " & vbCrLf
        sql += " ,dbo.DECODE10(cs.BudgetID,'03','就保','02','就安','01','公務','97','協助',cs.BudgetID) BudgetIDN" & vbCrLf
        '是主要參訓身分別為"一般身分者"、"農漁民"、"參加職業工會失業者"的學員，應繳自行負擔費用(元)
        '(自行負擔費用比率)
        sql += " ,case when cs.MIdentityID in ('01','11','17') then '20%' else '0%' end vbcRatio" & vbCrLf
        '(應繳自行負擔費用)
        sql += " ,case when cs.MIdentityID in ('01','11','17') then pp.defstdcost else 0 end defstdcost" & vbCrLf
        'sql += " ,pp.defstdcost" & vbCrLf
        '(receipt)收據號碼
        sql += " ,cs.receipt " & vbCrLf
        sql += " ,k1.Name MIName" & vbCrLf
        sql &= " FROM Class_ClassInfo cc" & vbCrLf
        sql &= " JOIN Plan_PlanInfo pp ON cc.PlanID = pp.PlanID AND cc.ComIDNO = pp.ComIDNO AND cc.SeqNO = pp.SeqNO" & vbCrLf
        'sql &= " JOIN id_class ic on ic.clsid =cc.clsid" & vbCrLf
        sql &= " join id_plan ip on ip.planid =cc.planid" & vbCrLf
        sql &= " join Key_TrainType tt on tt.TMID=cc.TMID" & vbCrLf
        sql &= " join Auth_Relship rr on rr.RID=cc.RID" & vbCrLf
        sql &= " join Org_OrgInfo oo on oo.comidno=cc.comidno" & vbCrLf
        sql += " join Class_StudentsOfClass cs on cs.ocid =cc.ocid   " & vbCrLf
        sql += " JOIN Stud_StudentInfo ss ON ss.SID=cs.SID " & vbCrLf
        sql += " JOIN Stud_SubData ss2 on ss2.SID =cs.SID " & vbCrLf
        sql += " LEFT JOIN Key_Identity k1 on k1.IdentityID =cs.MIdentityID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql += " and cs.StudStatus not in (2,3)" & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y'" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        sql &= " and ip.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " and ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql &= " and ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If
        'If RIDValue.Value <> "" Then
        '    sql &= " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
        'Else
        '    sql &= " and cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        'End If
        sql &= " and cc.OCID='" & vOCID & "'" & vbCrLf
        sql &= " order by StudID" & vbCrLf

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)
        labmsg.Text = "查無資料"
        DataGridS1.Visible = False
        btnSave1.Visible = False
        btnBack1.Visible = False
        If dt.Rows.Count > 0 Then
            'CPdt = dt.Copy()
            labmsg.Text = ""
            DataGridS1.Visible = True
            btnSave1.Visible = True
            btnBack1.Visible = True

            DataGridS1.DataSource = dt
            DataGridS1.DataBind()

            'PageControler1.PageDataTable = dt
            'PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGridC1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGridC1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "ocid")
        Dim vRID As String = TIMS.GetMyValue(sCmdArg, "rid")
        Select Case e.CommandName
            Case "add1" '新增
                'DataGridC1.Visible = False
                Call show_student(vOCID)
            Case "update1" '修改
                'DataGridC1.Visible = False
                Call show_student(vOCID)
            Case "print1" '列印
                'Common.MessageBox(Me, "列印")
                sCmdArg = ""
                TIMS.SetMyValue(sCmdArg, "OCID", vOCID)
                TIMS.SetMyValue(sCmdArg, "RID", vRID)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, sCmdArg)
                Exit Sub
        End Select
    End Sub

    Private Sub DataGridC1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGridC1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtAdd1 As LinkButton = e.Item.FindControl("lbtAdd1")
                Dim lbtUpdate1 As LinkButton = e.Item.FindControl("lbtUpdate1")
                Dim lbtPrint1 As LinkButton = e.Item.FindControl("lbtPrint1")

                lbtAdd1.Visible = True '未填寫採新增模式
                lbtUpdate1.Visible = False
                lbtPrint1.Visible = False

                If Chk_Receipt1(drv("ocid")) Then
                    '收據號碼
                    lbtAdd1.Visible = False
                    lbtUpdate1.Visible = True '填寫 可修改
                    lbtPrint1.Visible = True '填寫 可列印
                End If
                Dim cmdArg As String = ""
                cmdArg = ""
                TIMS.SetMyValue(cmdArg, "OCID", drv("ocid"))
                TIMS.SetMyValue(cmdArg, "RID", drv("rid"))

                lbtAdd1.CommandArgument = cmdArg
                lbtUpdate1.CommandArgument = cmdArg
                lbtPrint1.CommandArgument = cmdArg

        End Select

    End Sub

    '檢查是否有填寫
    Function Chk_Receipt1(ByVal ocid As String) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = ""
        sql = "select 'x' from CLASS_STUDENTSOFCLASS where ocid=@ocid and receipt is not null"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("ocid", SqlDbType.VarChar).Value = ocid
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    '回上頁
    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        DataGridS1.Visible = False
        btnSave1.Visible = False
        btnBack1.Visible = False

        tbSearch1.Visible = True
        DataGridC1.Visible = True
        PageControler1.Visible = True
        Call Search1()
    End Sub

    '存檔
    Sub savedata1()
        Dim sql As String = ""
        sql = ""
        sql &= " UPDATE CLASS_STUDENTSOFCLASS "
        sql &= " set receipt=@receipt" '收據號碼
        sql &= " where 1=1"
        sql &= " and socid =@socid"
        sql &= " and ocid =@ocid"
        Dim uCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        For Each eItem As DataGridItem In Me.DataGridS1.Items
            'Dim drv As DataRowView = eItem.DataItem
            Dim hidsocid As HiddenField = eItem.FindControl("hidsocid")
            Dim receipt As TextBox = eItem.FindControl("receipt")
            'hidsoicd.Value = Convert.ToString(drv("socid"))
            'receipt.Text = Convert.ToString(drv("receipt"))
            If hidsocid.Value <> "" Then
                receipt.Text = TIMS.ClearSQM(receipt.Text)
                With uCmd
                    .Parameters.Clear()
                    If receipt.Text <> "" Then
                        .Parameters.Add("receipt", SqlDbType.VarChar).Value = receipt.Text
                    Else
                        .Parameters.Add("receipt", SqlDbType.VarChar).Value = Convert.DBNull
                    End If
                    .Parameters.Add("socid", SqlDbType.VarChar).Value = hidsocid.Value
                    .Parameters.Add("ocid", SqlDbType.VarChar).Value = Hidocid.Value
                    .ExecuteNonQuery()
                End With
            End If
        Next

    End Sub

    '存檔
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Call savedata1()
        Common.MessageBox(Me, "儲存完成!!")

        DataGridS1.Visible = False
        btnSave1.Visible = False
        btnBack1.Visible = False

        tbSearch1.Visible = True
        DataGridC1.Visible = True
        PageControler1.Visible = True
        Call Search1()

    End Sub

    Private Sub DataGridS1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGridS1.ItemDataBound
        '<asp@HiddenField ID="hidsoicd" runat="server" />
        '<asp@TextBox ID="receipt"
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim hidsocid As HiddenField = e.Item.FindControl("hidsocid")
                Dim receipt As TextBox = e.Item.FindControl("receipt")
                hidsocid.Value = Convert.ToString(drv("socid"))
                receipt.Text = Convert.ToString(drv("receipt"))

        End Select

    End Sub

    Protected Sub DataGridS1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGridS1.SelectedIndexChanged

    End Sub

End Class
