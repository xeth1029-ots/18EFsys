Public Class CR_01_001
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'SYS_GRADEQUOTA
    'SELECT TOP 10 * FROM SYS_GRADEQUOTA
    ',dbo.FN_SCORING2_UPLIMIT(oo.COMIDNO,@TPLANID,@YEARS,@APPSTAGE,@ORGKIND2) UPLIMIT --'可核配上限,等級額度核配上限
    ',dbo.FN_SCORING2_GRADE(oo.COMIDNO,@TPLANID,@YEARS,@APPSTAGE) GRADE --'跨4區確認等級 跨區等級

    'Const cst_SCORELEVEL_B As String = "B"
    'Const cst_SCORELEVEL_C As String = "C"
    'Const cst_SCORELEVEL_D As String = "D"
    'Dim s_SCORELEVEL_all As String = "A,B,C,D"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            CCreate1()
        End If
    End Sub

    Sub CCreate1()
        msg1.Text = ""
        tbDataGrid1.Visible = False

        ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")
    End Sub

    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Call sSearch1()
    End Sub

    Sub sSearch1()
        msg1.Text = TIMS.cst_NODATAMsg1
        tbDataGrid1.Visible = False
        'TIMS.cst_SCORELEVEL_all 'As String = "A,B,C,D" '審核等級
        Dim A_SCORELEVEL_all As String() = TIMS.cst_SCORELEVEL_all.Split(",")

        Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH)
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        If v_YEARS_SCH = "" OrElse v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return
        End If
        Dim v_YEARS_SCH_ROC As String = TIMS.GET_YEARS_ROC(v_YEARS_SCH)

        Dim dt1 As New DataTable
        Dim dr1 As DataRow = Nothing
        dt1.Columns.Add(New DataColumn("YEARS"))
        dt1.Columns.Add(New DataColumn("YEARS_ROC"))
        dt1.Columns.Add(New DataColumn("APPSTAGE"))
        dt1.Columns.Add(New DataColumn("APPSTAGE_N"))
        dt1.Columns.Add(New DataColumn("SCORELEVEL"))
        dt1.Columns.Add(New DataColumn("CLASSQUOTAG"))
        dt1.Columns.Add(New DataColumn("CLASSQUOTAW"))

        For Each v_SCORELEVEL As String In A_SCORELEVEL_all
            dr1 = dt1.NewRow
            dt1.Rows.Add(dr1)
            Dim s_CLASSQUOTAG As String = Get_CLASSQUOTA(v_YEARS_SCH, v_APPSTAGE_SCH, v_SCORELEVEL, "G")
            Dim s_CLASSQUOTAW As String = Get_CLASSQUOTA(v_YEARS_SCH, v_APPSTAGE_SCH, v_SCORELEVEL, "W")
            dr1("YEARS") = v_YEARS_SCH
            dr1("YEARS_ROC") = v_YEARS_SCH_ROC 'TIMS.GET_YEARS_ROC(v_YEARS_SCH)
            dr1("APPSTAGE") = v_APPSTAGE_SCH
            dr1("APPSTAGE_N") = TIMS.Get_APPSTAGE2_N(v_APPSTAGE_SCH)
            dr1("SCORELEVEL") = v_SCORELEVEL
            dr1("CLASSQUOTAG") = s_CLASSQUOTAG
            dr1("CLASSQUOTAW") = s_CLASSQUOTAW
        Next

        msg1.Text = ""
        tbDataGrid1.Visible = True
        DataGrid1.DataSource = dt1
        DataGrid1.DataBind()
    End Sub

    ''' <summary> 取得 班級可核定額度 KINDGW - G:產投 /W:自主</summary>
    ''' <param name="v_YEARS"></param>
    ''' <param name="v_APPSTAGE"></param>
    ''' <param name="v_SCORELEVEL"></param>
    ''' <param name="v_KINDGW"></param>
    ''' <returns></returns>
    Function Get_CLASSQUOTA(ByRef v_YEARS As String, ByRef v_APPSTAGE As String, ByRef v_SCORELEVEL As String, ByRef v_KINDGW As String) As String
        Dim rst As String = ""
        'ByRef Htb As Hashtable
        'Dim v_YEARS As String = TIMS.GetMyValue2(Htb, "YEARS")
        'Dim v_APPSTAGE As String = TIMS.GetMyValue2(Htb, "APPSTAGE")
        'Dim v_SCORELEVEL As String = TIMS.GetMyValue2(Htb, "SCORELEVEL")

        Dim parms As New Hashtable From {{"YEARS", v_YEARS}, {"APPSTAGE", v_APPSTAGE}, {"SCORELEVEL", v_SCORELEVEL}}

        Dim sql As String = ""
        sql &= " SELECT a.SGQID" & vbCrLf 'sql &= "   /*PK*/" & vbCrLf
        sql &= " ,a.YEARS,a.APPSTAGE" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,a.SCORELEVEL" & vbCrLf
        sql &= " ,a.CLASSQUOTAG" & vbCrLf
        sql &= " ,a.CLASSQUOTAW" & vbCrLf
        sql &= " FROM SYS_GRADEQUOTA a" & vbCrLf
        sql &= " WHERE a.YEARS=@YEARS" & vbCrLf
        sql &= " and a.APPSTAGE=@APPSTAGE" & vbCrLf
        sql &= " and a.SCORELEVEL=@SCORELEVEL" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt Is Nothing Then Return rst
        If dt.Rows.Count = 0 Then Return rst

        Select Case v_KINDGW
            Case "G"
                rst = Convert.ToString(dt.Rows(0)("CLASSQUOTAG"))
            Case "W"
                rst = Convert.ToString(dt.Rows(0)("CLASSQUOTAW"))
        End Select
        Return rst
    End Function

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim Hid_YEARS As HiddenField = e.Item.FindControl("Hid_YEARS")
                Dim Hid_APPSTAGE As HiddenField = e.Item.FindControl("Hid_APPSTAGE")
                Dim Hid_SCORELEVEL As HiddenField = e.Item.FindControl("Hid_SCORELEVEL")
                Dim txtCLASSQUOTAG As TextBox = e.Item.FindControl("txtCLASSQUOTAG")
                Dim txtCLASSQUOTAW As TextBox = e.Item.FindControl("txtCLASSQUOTAW")
                Dim Hid_CLASSQUOTAG As HiddenField = e.Item.FindControl("Hid_CLASSQUOTAG")
                Dim Hid_CLASSQUOTAW As HiddenField = e.Item.FindControl("Hid_CLASSQUOTAW")

                'Dim sCmdArg As String = ""
                'TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                'TIMS.SetMyValue(sCmdArg, "APPSTAGE", Convert.ToString(drv("APPSTAGE")))
                'TIMS.SetMyValue(sCmdArg, "SCORELEVEL", Convert.ToString(drv("SCORELEVEL")))
                'TIMS.SetMyValue(sCmdArg, "CLASSQUOTA", Convert.ToString(drv("CLASSQUOTA")))
                Hid_YEARS.Value = Convert.ToString(drv("YEARS"))
                Hid_APPSTAGE.Value = Convert.ToString(drv("APPSTAGE"))
                Hid_SCORELEVEL.Value = Convert.ToString(drv("SCORELEVEL"))
                txtCLASSQUOTAG.Text = Convert.ToString(drv("CLASSQUOTAG"))
                txtCLASSQUOTAW.Text = Convert.ToString(drv("CLASSQUOTAW"))
                Hid_CLASSQUOTAG.Value = Convert.ToString(drv("CLASSQUOTAG"))
                Hid_CLASSQUOTAW.Value = Convert.ToString(drv("CLASSQUOTAW"))
        End Select
    End Sub

    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Call sSaveData1()
    End Sub

    Sub sSaveData1()
        Dim s_sql As String = ""
        s_sql &= " SELECT SGQID FROM SYS_GRADEQUOTA" & vbCrLf
        s_sql &= " WHERE YEARS=@YEARS AND APPSTAGE=@APPSTAGE AND SCORELEVEL=@SCORELEVEL" & vbCrLf
        'Dim sCmd As New SqlCommand(s_sql, objconn)

        Dim i_sql As String = ""
        i_sql &= " INSERT INTO SYS_GRADEQUOTA(SGQID ,YEARS ,APPSTAGE ,SCORELEVEL ,CLASSQUOTAG,CLASSQUOTAW ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@SGQID ,@YEARS ,@APPSTAGE ,@SCORELEVEL ,@CLASSQUOTAG,@CLASSQUOTAW ,@MODIFYACCT ,GETDATE())" & vbCrLf
        'Dim iCmd As New SqlCommand(i_sql, objconn)

        Dim u_sql As String = ""
        u_sql &= " UPDATE SYS_GRADEQUOTA" & vbCrLf
        u_sql &= " SET CLASSQUOTAG=@CLASSQUOTAG" & vbCrLf
        u_sql &= " ,CLASSQUOTAW=@CLASSQUOTAW" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE SGQID=@SGQID" & vbCrLf
        u_sql &= " AND SCORELEVEL=@SCORELEVEL" & vbCrLf
        'Dim uCmd As New SqlCommand(u_sql, objconn)

        Dim iRst As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim Hid_YEARS As HiddenField = eItem.FindControl("Hid_YEARS")
            Dim Hid_APPSTAGE As HiddenField = eItem.FindControl("Hid_APPSTAGE")
            Dim Hid_SCORELEVEL As HiddenField = eItem.FindControl("Hid_SCORELEVEL")
            Dim txtCLASSQUOTAG As TextBox = eItem.FindControl("txtCLASSQUOTAG")
            Dim txtCLASSQUOTAW As TextBox = eItem.FindControl("txtCLASSQUOTAW")
            Dim Hid_CLASSQUOTAG As HiddenField = eItem.FindControl("Hid_CLASSQUOTAG")
            Dim Hid_CLASSQUOTAW As HiddenField = eItem.FindControl("Hid_CLASSQUOTAW")

            '立即將控制權轉移到迴圈的下一個反復專案。
            Dim flag_Continue As Boolean = False
            If txtCLASSQUOTAG.Text = "" AndAlso txtCLASSQUOTAW.Text = "" Then flag_Continue = True '都沒有值，不儲存
            If Hid_CLASSQUOTAG.Value <> "" AndAlso Hid_CLASSQUOTAW.Value <> "" Then
                If Hid_CLASSQUOTAG.Value = txtCLASSQUOTAG.Text AndAlso Hid_CLASSQUOTAW.Value = txtCLASSQUOTAW.Text Then
                    flag_Continue = True '由於有值，且都沒有異動，所以不儲存
                End If
            End If
            '立即將控制權轉移到迴圈的下一個反復專案。
            If flag_Continue Then Continue For

            Dim parms As New Hashtable From {
                {"YEARS", Hid_YEARS.Value},
                {"APPSTAGE", Hid_APPSTAGE.Value},
                {"SCORELEVEL", Hid_SCORELEVEL.Value}
            }
            Dim dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, parms)

            If dt.Rows.Count = 0 Then
                Dim iSGQID As Integer = DbAccess.GetNewId(objconn, "SYS_GRADEQUOTA_SGQID_SEQ,SYS_GRADEQUOTA,SGQID")
                Dim i_parms As New Hashtable From {
                    {"SGQID", iSGQID},
                    {"YEARS", Hid_YEARS.Value},
                    {"APPSTAGE", Hid_APPSTAGE.Value},
                    {"SCORELEVEL", Hid_SCORELEVEL.Value},
                    {"CLASSQUOTAG", If(txtCLASSQUOTAG.Text <> "", Val(txtCLASSQUOTAG.Text), Convert.DBNull)},
                    {"CLASSQUOTAW", If(txtCLASSQUOTAW.Text <> "", Val(txtCLASSQUOTAW.Text), Convert.DBNull)}, 'Val(txtCLASSQUOTAW.Text)
                    {"MODIFYACCT", sm.UserInfo.UserID}
                }
                iRst += DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
            Else
                Dim dr1 As DataRow = dt.Rows(0)
                Dim iSGQID As Integer = TIMS.CINT1(dr1("SGQID"))
                Dim u_parms As New Hashtable From {
                    {"CLASSQUOTAG", If(txtCLASSQUOTAG.Text <> "", Val(txtCLASSQUOTAG.Text), Convert.DBNull)},
                    {"CLASSQUOTAW", If(txtCLASSQUOTAW.Text <> "", Val(txtCLASSQUOTAW.Text), Convert.DBNull)}, 'Val(txtCLASSQUOTAW.Text)
                    {"MODIFYACCT", sm.UserInfo.UserID},
                    {"SGQID", iSGQID},
                    {"SCORELEVEL", Hid_SCORELEVEL.Value}
                }
                iRst += DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
            End If
        Next
        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
    End Sub

End Class