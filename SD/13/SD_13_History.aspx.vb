Partial Class SD_13_History
    Inherits AuthBasePage

    Dim center_v As String = ""
    Dim RIDValue_v As String = ""
    'Const cst_printASPX_Q As String = "SD_14_002_Q.aspx?ID=" 'OLD
    Const cst_printASPX_R As String = "../14/SD_14_002_R.aspx?ID=" 'NEW
    'Type: A:已轉班查詢 B:未轉班查詢
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        center_v = sm.UserInfo.OrgName
        RIDValue_v = sm.UserInfo.RID
        Hid_ROCYEARS.Value = Convert.ToString(sm.UserInfo.Years - 1911)

        If Not IsPostBack Then
            CeartData()
        End If
    End Sub

    Sub CeartData()
        DataGrid1.Visible = False
        msg.Visible = True
        msg.Text = "查無學員重複參訓資料!!"
        TIMS.Tooltip(msg, "開訓日1年內的班級資料")

        Dim rqOCID As String = Convert.ToString(Request("OCID"))
        rqOCID = TIMS.ClearSQM(rqOCID)
        If rqOCID Is Nothing Then Exit Sub
        Dim drCC As DataRow = TIMS.GetOCIDDate(rqOCID, objconn)
        If drCC Is Nothing Then Exit Sub

        'Dim sql As String
        Dim dt As DataTable = TIMS.GET_Duplicate_Student(rqOCID, 2, objconn) '是否有重複參訓學員不排除產學訓(開訓日1年內的班級資料)
        If dt Is Nothing Then Exit Sub
        If dt.Rows.Count = 0 Then Exit Sub

        DataGrid1.Visible = True
        msg.Visible = False
        msg.Text = ""

        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub


    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "" Then Return ' Exit Sub
        If e.CommandArgument = "" Then Return ' Exit Sub
        If (sm.UserInfo.LID = 2) Then Return '(委訓單位不可使用)

        Dim sUrl As String = $"{cst_printASPX_R}{TIMS.Get_MRqID(Me)}"

        Hid_OCID1.Value = TIMS.GetMyValue(e.CommandArgument, "OCID")
        If Hid_OCID1.Value = "" Then Return ' Exit Sub'(查無班級不可使用)

        Dim xBlockN As String = TIMS.GetMyValue(e.CommandArgument, "xBlockN")
        If xBlockN = "" Then Return ' Exit Sub'(查無另開視窗序號不可使用)

        Dim drCC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        If drCC Is Nothing Then Return ' Exit Sub'(查無班級不可使用)

        Hid_ROCYEARS.Value = TIMS.ClearSQM(Hid_ROCYEARS.Value)
        Dim sCmdArg As String = ""
        TIMS.SetMyValue(sCmdArg, "Type", "A") 'Type: A:已轉班查詢 B:未轉班查詢
        TIMS.SetMyValue(sCmdArg, "PrintOrg", "Y") '顯示訓練單位名稱
        TIMS.SetMyValue(sCmdArg, "Years", Hid_ROCYEARS.Value)
        TIMS.SetMyValue(sCmdArg, "OCID", drCC("OCID"))
        TIMS.SetMyValue(sCmdArg, "FTYPE", "2") '1:細明體/2:標楷體(def)
        TIMS.SetMyValue(sCmdArg, "MSD", drCC("MSD")) '1:細明體/2:標楷體(def)

        Select Case e.CommandName
            Case "Link1"
                Dim url1 As String = String.Concat(sUrl, sCmdArg)
                Call TIMS.OpenWin1(Me, url1, xBlockN)
                'Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "Link2"
                Dim url1 As String = String.Concat(sUrl, sCmdArg)
                Call TIMS.OpenWin1(Me, url1, xBlockN)
                'Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drV As DataRowView = e.Item.DataItem 'drV
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                Dim lab_Classname As Label = e.Item.FindControl("lab_Classname")
                Dim lab_Classname2 As Label = e.Item.FindControl("lab_Classname2")
                Dim lib_Classname As LinkButton = e.Item.FindControl("lib_Classname")
                Dim lib_Classname2 As LinkButton = e.Item.FindControl("lib_Classname2")

                lib_Classname.Visible = False
                lib_Classname2.Visible = False
                lab_Classname.Visible = False
                lab_Classname2.Visible = False
                If (sm.UserInfo.LID = 2) Then
                    lab_Classname.Text = Convert.ToString(drV("Classname"))
                    lab_Classname2.Text = Convert.ToString(drV("Classname2"))
                    lab_Classname.Visible = True
                    lab_Classname2.Visible = True
                Else
                    lib_Classname.Text = Convert.ToString(drV("Classname")) 'Link1
                    lib_Classname2.Text = Convert.ToString(drV("Classname2")) 'Link2
                    lib_Classname.Visible = True
                    lib_Classname2.Visible = True

                    Const cst_t1 As String = "開啟訓練班別計畫表，查看上課時間"
                    Dim xBlockN As String = String.Concat("_MSG_", drV("OCID"), "_", Now.ToString("fffss")) '另開視窗序號
                    Dim sCmdArg As String = ""
                    TIMS.SetMyValue(sCmdArg, "OCID", drV("OCID"))
                    TIMS.SetMyValue(sCmdArg, "xBlockN", xBlockN)
                    lib_Classname.CommandArgument = sCmdArg
                    TIMS.Tooltip(lib_Classname, cst_t1)

                    Dim sCmdArg2 As String = ""
                    Dim xBlockN2 As String = String.Concat("_MSG_", drV("OCID2"), "_", Now.ToString("fffss")) '另開視窗序號
                    TIMS.SetMyValue(sCmdArg2, "OCID", drV("OCID2"))
                    TIMS.SetMyValue(sCmdArg2, "xBlockN", xBlockN2)
                    lib_Classname2.CommandArgument = sCmdArg2
                    TIMS.Tooltip(lib_Classname, cst_t1)
                End If
        End Select
    End Sub

End Class
