Public Class SD_14_020
    Inherits AuthBasePage

    Const cst_print_FN1 As String = "SD_14_020"

    Const cst_Plan As String = "Plan"
    Const cst_Class As String = "Class"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        Years.Value = sm.UserInfo.Years - 1911
        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            ClassTR.Visible = False

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'Me.Radio1.SelectedIndex = 0
            Common.SetListItem(rblClassType1, "0")

            PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "1")

            '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
            If tr_AppStage_TP28.Visible Then
                AppStage2 = TIMS.Get_APPSTAGE2(AppStage2)
                TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, AppStage2) 'Common.SetListItem(AppStage2, "3")
            End If

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2
            '列印
            'btnPrint1.Attributes("onclick") = "return CheckPrint();"
            '清除
            Button4.Attributes("onclick") = "ClearData();"
            'AllPrint.Visible = False
            'AllPrint.Attributes("onclick") = "SelectAll3(this.checked);"
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        Dim v_rblClassType1 As String = TIMS.GetListValue(rblClassType1)
        ClassTR.Visible = False
        If v_rblClassType1 <> "0" Then
            ClassTR.Visible = True
            TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
            If HistoryTable.Rows.Count <> 0 Then
                OCID1.Attributes("onclick") = "showObj('HistoryList');"
                OCID1.Style("CURSOR") = "hand"
            End If
        End If

    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim v_rblClassType1 As String = TIMS.GetListValue(rblClassType1)
        Select Case v_rblClassType1
            Case "0"
                Call CreateClass(cst_Plan)
            Case Else
                Call CreateClass(cst_Class)
        End Select

    End Sub

    Sub CreateClass(ByVal sType As String)

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = Session("RID")
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim v_AppStage2 As String = "" 'TIMS.GetListValue(AppStage2)
        If (v_AppStage2 <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage2
        If tr_AppStage_TP28.Visible Then v_AppStage2 = TIMS.GetListValue(AppStage2)

        Dim SQL_1 As String = "
SELECT CC.OCID,PP.PLANID,PP.COMIDNO,PP.SEQNO
,CONCAT(dbo.FN_GET_CLASSCNAME(PP.CLASSNAME,PP.CYCLTYPE),CASE WHEN PP.RESULTBUTTON IN ('Y','R') THEN '(未送出)' END) CLASSCNAME
,format(ISNULL(cc.STDate,pp.STDate),'yyyy/MM/dd') STDATE
,format(ISNULL(cc.FTDate,pp.FDDate),'yyyy/MM/dd') FTDATE
,b.ORGNAME
FROM PLAN_PLANINFO PP
JOIN VIEW_RIDNAME B ON B.RID=PP.RID
JOIN ID_PLAN IP ON IP.PLANID=PP.PLANID
LEFT JOIN CLASS_CLASSINFO CC ON CC.PLANID=PP.PLANID AND CC.COMIDNO=PP.COMIDNO AND CC.SEQNO=PP.SEQNO
WHERE PP.IsApprPaper='Y'
"
        SQL_1 &= $" AND ip.TPlanID='{sm.UserInfo.TPlanID}' AND ip.YEARS='{sm.UserInfo.Years}'" & vbCrLf
        If sm.UserInfo.LID <> 0 Then
            SQL_1 &= $" AND PP.PLANID={sm.UserInfo.PlanID}" & vbCrLf
        End If
        SQL_1 &= " AND B.RELSHIP like '" & RelShip & "%'" & vbCrLf
        Select Case sType
            Case cst_Plan '未轉班
                SQL_1 &= " AND pp.TransFlag='N'" & vbCrLf
            Case cst_Class '已轉班 '且班級存在
                SQL_1 &= " AND pp.TransFlag='Y' AND cc.OCID IS NOT NULL" & vbCrLf
                If OCIDValue1.Value <> "" Then
                    SQL_1 &= " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf '若有選擇班級
                End If
        End Select
        '依申請階段
        If v_AppStage2 <> "" Then SQL_1 &= " AND pp.AppStage='" & v_AppStage2 & "'" & vbCrLf

        '28:產業人才投資方案
        orgid.Value = ""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                '產業人才投資計畫
                SQL_1 &= " AND b.OrgKind <> '10'" & vbCrLf
                orgid.Value = "G"
            Else
                '提升勞工自主學習計畫
                SQL_1 &= " AND b.OrgKind = '10'" & vbCrLf
                orgid.Value = "W"
            End If
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(SQL_1, objconn)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        DataGrid1.Visible = False
        'PageControler1.Visible = False
        'DataGrid2.Visible = False
        'PageControler2.Visible = False
        If TIMS.dtNODATA(dt) Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        DataGrid1.Visible = True
        'PageControler1.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub


    Sub UTL_PRINT1(ByRef s_prt_val As String)
        Dim v_rblClassType1 As String = TIMS.GetListValue(rblClassType1)

        PLANIDValue.Value = TIMS.GetMyValue(s_prt_val, "PlanID")
        ComIDNOValue.Value = TIMS.GetMyValue(s_prt_val, "ComIDNO")
        SeqNoValue.Value = TIMS.GetMyValue(s_prt_val, "SeqNo")
        PCSValue.Value = TIMS.GetMyValue(s_prt_val, "PCS")
        OCIDValue.Value = TIMS.GetMyValue(s_prt_val, "OCID")

        Dim MyValue As String = ""
        MyValue = "YEARS=" & TIMS.ClearSQM(Years.Value)
        MyValue += "&PLANID=" & TIMS.ClearSQM(PLANIDValue.Value)
        MyValue += "&ComIDNO=" & TIMS.ClearSQM(ComIDNOValue.Value)
        MyValue += "&SEQNO=" & TIMS.ClearSQM(SeqNoValue.Value)
        MyValue += "&PCSValue=" & TIMS.ClearSQM(PCSValue.Value)

        Select Case v_rblClassType1
            Case "0"
            Case Else
                MyValue += "&OCID=" & TIMS.ClearSQM(OCIDValue.Value)
        End Select

        '材料明細表
        'Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "BussinessTrain", "SD_14_020", MyValue)
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_print_FN1, MyValue)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument

        Select Case e.CommandName
            Case "Print1"
                UTL_PRINT1(sCmdArg)
        End Select
    End Sub


    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim chkbox1 As HtmlInputCheckBox = e.Item.FindControl("chkbox1")
                Dim hidOCID As HtmlInputHidden = e.Item.FindControl("hidOCID")
                Dim hidPlanID As HtmlInputHidden = e.Item.FindControl("hidPlanID")
                Dim hidComIDNO As HtmlInputHidden = e.Item.FindControl("hidComIDNO")
                Dim hidSeqNo As HtmlInputHidden = e.Item.FindControl("hidSeqNo")
                hidOCID.Value = Convert.ToString(drv("OCID"))
                hidPlanID.Value = Convert.ToString(drv("PlanID"))
                hidComIDNO.Value = Convert.ToString(drv("ComIDNO"))
                hidSeqNo.Value = Convert.ToString(drv("SeqNo"))
                Dim TMPPCS As String = Convert.ToString(drv("PlanID")) & "x" & Convert.ToString(drv("ComIDNO")) & "x" & Convert.ToString(drv("SeqNo"))

                Dim btnPrint1 As Button = e.Item.FindControl("btnPrint1")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "ComIDNO", Convert.ToString(drv("ComIDNO")))
                TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))
                TIMS.SetMyValue(sCmdArg, "PCS", TMPPCS)
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                btnPrint1.CommandArgument = sCmdArg
        End Select
    End Sub

    Protected Sub rblClassType1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rblClassType1.SelectedIndexChanged
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
    End Sub

End Class
