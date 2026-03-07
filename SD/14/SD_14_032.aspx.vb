Partial Class SD_14_032
    Inherits AuthBasePage

    Const cst_printFN_G1 As String = "SD_14_032G"
    Const cst_printFN_W1 As String = "SD_14_032W"

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1
        Years.Value = sm.UserInfo.Years - 1911

        If Not IsPostBack Then
            CCreate1()
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

        Button1.Attributes("onclick") = "return CheckSearch();"
    End Sub
    Sub CCreate1()
        msg.Text = ""
        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objconn)
        Common.SetListItem(PlanPoint, "0")

        If sm.UserInfo.LID <> "2" Then
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Hid_OrgKind2.Value = Convert.ToString(drCC("ORGKINDGW"))
        Hid_MSD.Value = Convert.ToString(drCC("MSD"))
        OCIDValue1.Value = Convert.ToString(drCC("OCID"))
        SOCIDValue.Value = ""

        Dim pParms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"OCID", OCIDValue1.Value}}

        Dim sSql As String = ""
        sSql &= " SELECT a.OCID,a.SOCID,a.StudentID,a.SID" & vbCrLf
        sSql &= " ,a.STUDID2,a.CLASSCNAME2,a.STDate,a.FTDate" & vbCrLf
        sSql &= " ,a.NAME,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,format(a.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK2(a.Birthday) BIRTHDAY_MK" & vbCrLf
        sSql &= " ,a.SEX_N,a.STUDSTATUS_N,a.Sex,a.StudStatus,a.AppliedResultM ,a.PlanID ,a.RID" & vbCrLf
        sSql &= " FROM dbo.VIEW_STUDENTBASICDATA a" & vbCrLf
        '排除離退訓學員
        sSql &= " WHERE a.STUDSTATUS NOT IN (2,3)" & vbCrLf
        sSql &= " AND a.TPLANID=@TPLANID AND a.YEARS=@YEARS" & vbCrLf
        sSql &= " AND a.OCID=@OCID" & vbCrLf

        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    sSql &= " AND a.OrgKind2 = 'G'" & vbCrLf
                Case "2"
                    '提升勞工自主學習計畫
                    sSql &= " AND a.OrgKind2 = 'W'" & vbCrLf
            End Select
        End If
        sSql &= " ORDER BY a.StudentID " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        msg.Text = "查無資料"
        DataGridTable.Visible = False

        If TIMS.dtNODATA(dt) Then Return '(沒資料就算了)

        msg.Text = ""
        DataGridTable.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
        'Call TIMS.set_row_color(DataGrid1)
    End Sub

    ''' <summary>比對字串後，如果有回應true 沒有為false</summary>
    ''' <param name="s_SOCID"></param>
    ''' <param name="s_SOCIDValue"></param>
    ''' <returns></returns>
    Function Comp1(ByRef s_SOCID As String, ByRef s_SOCIDValue As String) As Boolean
        If s_SOCID = "" OrElse s_SOCIDValue = "" Then Return False
        Dim SOCIDArray As String() = Split(s_SOCIDValue, ",")
        If SOCIDArray.Length = 0 Then Return False
        For Each str1 As String In SOCIDArray
            If str1.Equals(s_SOCID) Then Return True
        Next
        Return False
    End Function
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim SOCID As HtmlInputCheckBox = e.Item.FindControl("SOCID")
                SOCID.Value = Convert.ToString(drv("SOCID"))
                SOCID.Attributes("onclick") = "SelectItem(this.checked,this.value);"
                SOCID.Checked = Comp1(SOCID.Value, SOCIDValue.Value)
        End Select
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
    End Sub
    Function GET_DataGrid1_SOCIDVAL() As String
        Dim tSOCIDVal As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim SOCID As HtmlInputCheckBox = eItem.FindControl("SOCID")
            If SOCID.Value <> "" AndAlso SOCID.Checked Then
                tSOCIDVal &= String.Concat(If(tSOCIDVal <> "", ",", ""), TIMS.ClearSQM(SOCID.Value))
            End If
        Next
        Return tSOCIDVal
    End Function
    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
        '列印
        SOCIDValue.Value = GET_DataGrid1_SOCIDVAL()
        SOCIDValue.Value = TIMS.CombiSQM2IN(SOCIDValue.Value)
        SOCIDValue.Value = TIMS.ClearSQM(SOCIDValue.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Years.Value = TIMS.ClearSQM(Years.Value)
        If SOCIDValue.Value = "" Then
            Common.MessageBox(Me, "請選擇要列印的學員")
            Exit Sub
        End If
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim myValue As String = ""
        TIMS.SetMyValue(myValue, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(myValue, "MSD", Hid_MSD.Value)
        TIMS.SetMyValue(myValue, "RID", RIDValue.Value)
        TIMS.SetMyValue(myValue, "SOCID", SOCIDValue.Value)
        TIMS.SetMyValue(myValue, "Years", Years.Value)

        'SD_14_032G,,SD_14_032W 收據
        Hid_OrgKind2.Value = TIMS.ClearSQM(Hid_OrgKind2.Value)
        Dim prtFilename As String = ""
        Select Case Hid_OrgKind2.Value 'hidorgid.Value
            Case "G"
                prtFilename = cst_printFN_G1
            Case "W"
                prtFilename = cst_printFN_W1
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, myValue)

    End Sub
End Class
