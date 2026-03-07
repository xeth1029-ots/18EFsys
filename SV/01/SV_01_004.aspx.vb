Partial Class SV_01_004
    Inherits AuthBasePage

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

        '分頁設定

        PageControler1.PageDataGrid = DataGrid1
        Pagecontroler2.PageDataGrid = DataGrid2
        Pagecontroler3.PageDataGrid = DataGrid3

        If Not IsPostBack Then

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            PageControler1.Visible = False
            Pagecontroler2.Visible = False
            Pagecontroler3.Visible = False
            table_Q.Visible = True
            DataGrid1Table.Visible = False
            Classtable.Visible = False
            DataGrid2table.Visible = False
            DataGrid3table.Visible = False

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button7_Click(sender, e)
            End If

            Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
            If rqOCID <> "" Then
                Ipt_Name.Value = TIMS.ClearSQM(Request("IptName"))
                SVID.Value = TIMS.ClearSQM(Request("SVID"))
                RIDValue.Value = TIMS.ClearSQM(Request("RID"))
                OCIDValue1.Value = TIMS.ClearSQM(Request("OCIDValue1"))
                'PageControler3.PageDataGrid = DataGrid3
                Pagecontroler3.PageIndex = Val(TIMS.ClearSQM(Request("PG")))
                Call Search_Student(rqOCID)
            End If

        End If

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

    End Sub

    'Private Sub search_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.ServerClick
    '    dt_search() '查詢
    'End Sub

    Sub dt_search()
        'Dim str As String
        'Dim sql As String
        'Dim dt As DataTable

        Ipt_Name.Value = TIMS.ClearSQM(Ipt_Name.Value)
        Dim sql As String = ""
        sql = ""
        sql &= " select a.SVID"
        sql &= " ,a.Name"
        sql &= " ,case a.Avail when 'Y' then '啟用' else '不啟用' end Avail"
        sql += " ,a.Avail ISUSE" & vbCrLf
        sql += " ,a.internal " & vbCrLf
        sql &= " from ID_SURVEY a"
        sql &= " where 1=1"
        'sql &= " AND a.INTERNAL IS NULL"
        sql &= " and a.Avail <> 'N'"
        If Ipt_Name.Value <> "" Then   '搜尋條件
            sql &= " and a.Name like '%" & Ipt_Name.Value & "%' "
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        msg.Visible = True
        If dt.Rows.Count = 0 Then
            table_Q.Visible = True
            DataGrid1.Visible = False
            PageControler1.Visible = False
            Pagecontroler2.Visible = False
            Pagecontroler3.Visible = False
            Classtable.Visible = False
            DataGrid2table.Visible = False
            DataGrid3table.Visible = False
            Exit Sub
        End If
        msg.Text = ""
        'msg.Visible = True

        table_Q.Visible = True
        DataGrid1Table.Visible = True
        DataGrid1.Visible = True
        PageControler1.Visible = True
        Pagecontroler2.Visible = False
        Pagecontroler3.Visible = False
        Classtable.Visible = False
        DataGrid2table.Visible = False
        DataGrid3table.Visible = False
        'PageControler1.SqlString = sql
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim oDG1 As DataGrid = DataGrid1

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                'Dim SQL As String
                'Dim dt As DataTable
                Dim Btn_edit As Button = e.Item.FindControl("Btn_edit")
                Btn_edit.CommandArgument = CStr(drv("SVID")) 'e.Item.Cells(4).Text.ToString
                Dim dt As DataTable = Nothing
                dt = TIMS.Get_dtKSK(drv("SVID"), objconn)
                If dt.Rows.Count = 0 Then
                    Btn_edit.Enabled = False
                    Btn_edit.ToolTip = "尚未設定【問卷分類標題設定 】功能的問卷分類標題!!" & vbCrLf
                End If
                Dim iSVIDCnt As Integer = TIMS.GetSVIDCnt(drv("SVID"), objconn)
                If iSVIDCnt = 0 Then
                    Btn_edit.Enabled = False
                    Btn_edit.ToolTip += "尚未設定【問卷題目設定 】功能的問卷題目!! "
                End If

                'internal
                If Convert.ToString(drv("internal")) = "Y" Then
                    Btn_edit.CommandArgument = ""
                    Btn_edit.Enabled = False
                    Btn_edit.ToolTip = "內部使用!!"
                End If
        End Select

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim seCmdArg As String = e.CommandArgument
        If seCmdArg = "" Then Exit Sub

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                SVID.Value = seCmdArg
                Call Search_class()
        End Select

    End Sub

    Sub Search_class()

        table_Q.Visible = False
        DataGrid1Table.Visible = False
        PageControler1.Visible = False
        Classtable.Visible = True
        DataGrid2table.Visible = False
        Pagecontroler2.Visible = False
        Pagecontroler3.Visible = False
        DataGrid3table.Visible = False

    End Sub


    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGrid2table.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGrid2table.Visible = False
    End Sub


    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim seCmdArg As String = e.CommandArgument
        If seCmdArg = "" Then Exit Sub

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Call Search_Student(seCmdArg)
        End Select

    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        'Dim oDG2 As DataGrid = DataGrid2

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                Dim btnE As Button = e.Item.FindControl("EDIT")
                btnE.CommandArgument = CStr(drv("OCID")) 'e.Item.Cells(8).Text.ToString

        End Select
    End Sub


    Sub Search_Student(ByVal rqOCID As String)
        OCID2.Value = TIMS.ClearSQM(rqOCID)

        'Dim sql As String
        'Dim str3 As String
        'Dim dt As DataTable
        'OCID2.Value = OCID
        'If OCID.ToString <> "" Then
        '    str3 = " and cc.OCID = '" & OCID.ToString & "' "
        'End If

        Dim sql As String = ""
        sql = "select cs.SOCID,oo.orgname,cc.classcname + cc.cycltype as CName, " & vbCrLf
        sql += " CONVERT(varchar, cc.STdate, 111) + '~' + CONVERT(varchar, cc.FTdate, 111) as SEdate," & vbCrLf
        sql += " dbo.SUBSTR2(StudentID,-2) as StudentID,ss.Name as Sname, " & vbCrLf
        sql += " case when sy.Socid is not NULL then 1 else 0 end as isinput " & vbCrLf
        sql += " from org_orginfo oo join auth_relship ar on oo.orgid = ar.orgid" & vbCrLf
        sql += " join class_classinfo cc on ar.rid = cc.rid" & vbCrLf
        sql += " left join Class_StudentsOfClass cs on cc.ocid = cs.ocid " & vbCrLf
        sql += " join stud_studentinfo ss on cs.sid = ss.sid"
        sql += " left join (select DISTINCT SOCID from Stud_Survey where SVID = " & SVID.Value & ") sy on sy.Socid = cs.Socid " & vbCrLf
        sql += " where cc.notopen ='N'" & vbCrLf
        If rqOCID <> "" Then
            sql += " and cc.OCID = '" & rqOCID & "' "
        End If
        sql += " order by StudentID " & vbCrLf

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            msg3.Text = "查無資料"
            msg3.Visible = True
            table_Q.Visible = False
            DataGrid1Table.Visible = False
            DataGrid1.Visible = False
            PageControler1.Visible = False
            Pagecontroler2.Visible = False
            Pagecontroler3.Visible = False
            Classtable.Visible = False
            DataGrid2table.Visible = False
            DataGrid3table.Visible = True
            TR1.Visible = False
            TR2.Visible = False
            return2.Visible = False
        Else
            msg3.Visible = False
            table_Q.Visible = False
            DataGrid1Table.Visible = False
            DataGrid1.Visible = False
            PageControler1.Visible = False
            Pagecontroler2.Visible = True
            Pagecontroler3.Visible = False
            Classtable.Visible = False
            DataGrid2table.Visible = False
            ORGL2.Text = dt.Rows(0).Item(1).ToString
            CLASSL2.Text = dt.Rows(0).Item(2).ToString
            ODDATE2.Text = dt.Rows(0).Item(3).ToString
            DataGrid3table.Visible = True
            Pagecontroler3.Sort = "StudentID"
            'Pagecontroler3.SqlString = sql
            'Pagecontroler3.PrimaryKey = "SOCID"
            Pagecontroler3.PageDataTable = dt '.PrimaryKey = "SOCID"
            Pagecontroler3.ControlerLoad()
        End If
    End Sub

    Const Cst_是否有填問卷 As Integer = 4

    Private Sub DataGrid3_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim InsertBtn As Button = e.Item.FindControl("InsertBtn") '新增
                Dim EditBtn As Button = e.Item.FindControl("EditBtn") '編輯
                Dim DeleteBtn As Button = e.Item.FindControl("DeleteBtn") '刪除

                'Dim view As Button = e.Item.FindControl("view")
                InsertBtn.Visible = True
                EditBtn.Visible = True
                DeleteBtn.Visible = True
                DeleteBtn.Attributes("onclick") = TIMS.cst_confirm_delmsg1 '刪除

                If Convert.ToString(drv("isinput")) = "1" Then '判斷是否有填寫過問卷 1是有填過
                    InsertBtn.Enabled = False
                    EditBtn.Enabled = True
                    DeleteBtn.Enabled = True
                Else
                    InsertBtn.Enabled = True
                    EditBtn.Enabled = False
                    DeleteBtn.Enabled = False
                    EditBtn.Visible = False
                    DeleteBtn.Visible = False
                End If

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SOCID", drv("SOCID"))

                InsertBtn.CommandArgument = sCmdArg 'drv("SOCID") 'e.Item.Cells(3).Text.ToString
                EditBtn.CommandArgument = sCmdArg 'InsertBtn.CommandArgument
                DeleteBtn.CommandArgument = sCmdArg 'InsertBtn.CommandArgument
        End Select
    End Sub

    Private Sub DataGrid3_ItemCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid3.ItemCommand
        If e.CommandArgument = "" Then Return
        Dim sCmdArg As String = e.CommandArgument
        Dim SOCID_val As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        SVID.Value = TIMS.ClearSQM(SVID.Value)

        Select Case e.CommandName
            Case "E", "I"
                If SOCID_val = "" Then Return
                If SVID.Value = "" Then Return

                TIMS.Utl_Redirect1(Me, "~/AC/06/AC_06_002_Insert.aspx?ID=" & Request("ID") & "&SOCID=" & SOCID_val & "&SVID=" & SVID.Value & "&Type=" & e.CommandName & "&OCID=" & OCID2.Value & "&IptName=" & Ipt_Name.Value & "&RID=" & RIDValue.Value & "&OCIDValue1=" & OCIDValue1.Value & "&PG=" & DataGrid3.CurrentPageIndex + 1 & "")
            Case "D" '刪除
                'Del.Attributes("onclick") = TIMS.cst_confirm_delmsg1 '刪除
                If SOCID_val = "" Then Return
                If SVID.Value = "" Then Return

                Dim dParms As New Hashtable
                dParms.Add("SOCID", SOCID_val)
                dParms.Add("SVID", SVID)
                Dim d_Sql As String = " DELETE STUD_SURVEY where 1=1 AND SOCID =@SOCID AND SVID =@SVID"
                DbAccess.ExecuteNonQuery(d_Sql, objconn, dParms)

                Common.MessageBox(Me, "刪除成功")

                OCID2.Value = TIMS.ClearSQM(OCID2.Value)
                Call Search_Student(OCID2.Value)
        End Select

    End Sub

    '回上一頁
    Private Sub return1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles return1.Click
        PageControler1.Visible = False
        Pagecontroler2.Visible = False
        Pagecontroler3.Visible = False
        table_Q.Visible = True
        DataGrid1Table.Visible = False
        Classtable.Visible = False
        DataGrid2table.Visible = False
        DataGrid3table.Visible = False
        'search_ServerClick(sender, e)
        Call dt_search() '查詢
    End Sub

    '回上一頁
    Private Sub return2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles return2.Click
        table_Q.Visible = False
        DataGrid1Table.Visible = False
        PageControler1.Visible = False
        Classtable.Visible = True
        DataGrid2table.Visible = False
        Pagecontroler2.Visible = False
        Pagecontroler3.Visible = False
        DataGrid3table.Visible = False
        'Button1_Click(sender, e)
        Call sSearch2()
    End Sub

    '查詢
    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call dt_search() '查詢
    End Sub

    Protected Sub btnSearch2_Click(sender As Object, e As EventArgs) Handles btnSearch2.Click
        Call sSearch2()
    End Sub


    Sub sSearch2()
        'Dim str2 As String
        Dim sql As String

        'sql = "select cc.OCID,oo.orgname,cc.classcname + cc.cycltype as CName, " & vbCrLf
        'sql += " convert (varchar,cc.STdate,111) + '~' + convert(varchar,cc.FTdate,111) as SEdate," & vbCrLf
        'sql += " count(cs.Socid) as opencount," & vbCrLf
        'sql += " sum(dbo.NVL(case when cs.studstatus not in(2,3) and cc.FTdate < getdate() then 1 end,0)) as closecount," & vbCrLf
        'sql += " count(Distinct sy.socid) as inputcount" & vbCrLf
        'sql += " from org_orginfo oo join auth_relship ar on oo.orgid = ar.orgid" & vbCrLf
        'sql += " join class_classinfo cc on ar.rid = cc.rid" & vbCrLf
        'sql += " left join Class_StudentsOfClass cs on cc.ocid = cs.ocid " & vbCrLf
        'sql += " left join (select socid from Stud_survey where svid = " & SVID.Value & ") sy on cs.socid = sy.socid " & vbCrLf
        'sql += " where cc.notopen ='N' and cc.PlaniD = " & sm.UserInfo.PlanID & "" & vbCrLf
        'sql += " " & str2 & "" & vbCrLf
        'sql += " GROUP BY cc.OCID,oo.orgname,cc.classcname,cc.cycltype,cc.STdate,cc.FTdate "

        sql = "" & vbCrLf
        sql += " select " & vbCrLf
        sql += " 	cc.OCID,oo.orgname,cc.classcname + cc.cycltype as CName" & vbCrLf
        sql += " 	, CONVERT(varchar, cc.STdate, 111) + '~' + CONVERT(varchar, cc.FTdate, 111) as SEdate" & vbCrLf
        sql += " 	, count(cs.Socid) as opencount" & vbCrLf
        sql += " 	, sum(dbo.NVL(case when cs.studstatus not in(2,3) and cc.FTdate < getdate() then 1 end,0)) as closecount" & vbCrLf
        sql += " 	, count(Distinct sy.socid) as inputcount" & vbCrLf
        sql += " from " & vbCrLf
        sql += "  org_orginfo oo " & vbCrLf
        sql += "  join auth_relship ar on oo.orgid = ar.orgid" & vbCrLf
        sql += "  join class_classinfo cc on ar.rid = cc.rid" & vbCrLf
        sql += "  left join Class_StudentsOfClass cs on cc.ocid = cs.ocid " & vbCrLf
        sql += "  left join (select distinct socid from Stud_survey where svid =  " & SVID.Value & ") sy on cs.socid = sy.socid " & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and cc.notopen ='N'" & vbCrLf
        sql += " and cc.PlaniD =" & sm.UserInfo.PlanID & "" & vbCrLf
        'sql += "  " & str2 & "" & vbCrLf
        If RIDValue.Value <> "" Then
            sql += " and ar.RID = '" & RIDValue.Value & "' "
        End If
        If OCIDValue1.Value <> "" Then
            sql += " and cc.ocid = '" & OCIDValue1.Value & "'"
        End If
        sql += " GROUP BY " & vbCrLf
        sql += "  cc.OCID,oo.orgname,cc.classcname,cc.cycltype,cc.STdate,cc.FTdate " & vbCrLf


        If TIMS.Get_SQLRecordCount(sql, objconn) = 0 Then

            msg2.Text = "查無資料"
            msg2.Visible = True
            table_Q.Visible = False
            DataGrid1Table.Visible = False
            PageControler1.Visible = False
            Pagecontroler2.Visible = False
            Pagecontroler3.Visible = False
            Classtable.Visible = True
            DataGrid2.Visible = False
            return1.Visible = False
            DataGrid3table.Visible = False
            DataGrid2table.Visible = True
        Else
            msg2.Visible = False
            table_Q.Visible = False
            DataGrid1Table.Visible = False
            DataGrid1.Visible = False
            PageControler1.Visible = False
            Pagecontroler2.Visible = True
            Pagecontroler3.Visible = False
            Classtable.Visible = True
            DataGrid2.Visible = True
            DataGrid2table.Visible = True
            DataGrid3table.Visible = False
            Pagecontroler2.SqlString = sql
            Pagecontroler2.ControlerLoad()
        End If

    End Sub

End Class