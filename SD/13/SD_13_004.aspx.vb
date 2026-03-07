Partial Class SD_13_004
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            PageControler1.Visible = False
            DataGrid1.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'DataGrid1Table.Style("display") = "none"
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

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

        'If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
        '    Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');SetOneOCID();"
        'Else
        '    Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');SetOneOCID();"
        'End If
        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then '機構選擇
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If
        'Button1.Attributes("onclick") = "return CheckSearch();"
        'Button3.Attributes("onclick") = "return CheckData();"
    End Sub

    '查詢SQL
    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim sPMS As New Hashtable
        sPMS.Add("TPLANID", sm.UserInfo.TPlanID)
        sPMS.Add("YEARS", Convert.ToString(sm.UserInfo.Years))
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                sPMS.Add("DISTID", sm.UserInfo.DistID)
                sPMS.Add("PLANID", sm.UserInfo.PlanID)
        End Select
        If RIDValue.Value.Length > 1 Then sPMS.Add("RID", RIDValue.Value)
        If OCIDValue1.Value <> "" Then sPMS.Add("OCID", OCIDValue1.Value)

        Dim sSql As String = ""
        sSql &= " WITH WC1 AS ( SELECT cc.PLANID,cc.ORGNAME,cc.OCID,cc.STDATE,cc.FTDATE" & vbCrLf
        sSql &= " ,cc.CLASSCNAME2,cc.CLASSCNAME,cc.CYCLTYPE" & vbCrLf
        sSql &= " ,cc.TPLANID,cc.YEARS,cc.RID,cc.DISTID,cc.DISTNAME" & vbCrLf
        sSql &= " FROM VIEW2 cc" & vbCrLf
        sSql &= " WHERE cc.TPLANID =@TPLANID AND cc.YEARS =@YEARS" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                sSql &= " AND cc.DISTID=@DISTID" & vbCrLf
                sSql &= " AND cc.PLANID=@PLANID" & vbCrLf
        End Select
        If RIDValue.Value.Length > 1 Then sSql &= " AND cc.RID=@RID" & vbCrLf
        If OCIDValue1.Value <> "" Then sSql &= " AND cc.OCID=@OCID" & vbCrLf
        sSql &= " )" & vbCrLf
        sSql &= " ,WS1 AS ( SELECT cc.OCID" & vbCrLf
        sSql &= " 	,COUNT(CASE WHEN g.SOCID IS NOT NULL THEN 1 END) Num1" & vbCrLf
        sSql &= " 	,COUNT(CASE WHEN g.AppliedStatusM='Y' THEN 1 END) Num2" & vbCrLf
        sSql &= " 	,COUNT(CASE WHEN g.AppliedStatus='1' THEN 1 END) Num3" & vbCrLf
        sSql &= " 	FROM WC1 cc" & vbCrLf
        sSql &= " 	JOIN dbo.V_STUDENTINFO cs on cs.OCID=cc.OCID" & vbCrLf
        sSql &= " 	LEFT JOIN dbo.STUD_SUBSIDYCOST g ON g.SOCID = cs.SOCID" & vbCrLf
        sSql &= " 	GROUP BY cc.OCID )" & vbCrLf

        sSql &= " SELECT cc.ORGNAME,cc.OCID,cc.STDATE,cc.FTDATE" & vbCrLf
        sSql &= " ,cc.CLASSCNAME2,cc.CLASSCNAME,cc.CYCLTYPE" & vbCrLf
        sSql &= " ,cc.TPLANID,cc.YEARS,cc.RID,cc.DISTID,cc.DISTNAME" & vbCrLf
        sSql &= " ,cs.Num1,cs.Num2,cs.Num3" & vbCrLf
        sSql &= " FROM WC1 cc" & vbCrLf
        sSql &= " JOIN WS1 cs on cs.OCID=cc.OCID" & vbCrLf
        sSql &= " ORDER by cc.ORGNAME ,cc.CLASSCNAME2 ,cc.CYCLTYPE" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)

        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        msg.Text = ""
        PageControler1.Visible = True
        DataGrid1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.EditItem, ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
        End Select
    End Sub
End Class