Partial Class TC_01_017_add
    Inherits AuthBasePage

#Region "Sub"
    '代入DropDownList資料
    'Private Sub get_ddl_typeid2(ByVal TypeID1 As String)
    '    If TypeID1 = "" Or TypeID1 = "0" Then Exit Sub
    '    Dim sql As String = ""
    '    Dim dt As New DataTable
    '    sql = "" & vbCrLf
    '    sql &= " select TypeID1,TypeID2,concat(TypeID2,'-',TypeID2Name )TypeID2Name " & vbCrLf
    '    sql &= " FROM Key_OrgType1 "
    '    sql &= " WHERE TypeID1=@TypeID1"
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim oCmd As New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("@TypeID1", SqlDbType.Int).Value = CInt(TypeID1)
    '        dt.Load(.ExecuteReader())
    '    End With


    '    'Dim da As New SqlDataAdapter
    '    'Dim dt As New DataTable
    '    'Try
    '    '    sql = "select TypeID1,TypeID2,TypeID2 + '-' + TypeID2Name TypeID2Name from Key_OrgType1 "
    '    '    sql += "where TypeID1=@TypeID1"
    '    '    With da
    '    '        .SelectCommand = New SqlCommand(sql, objconn)
    '    '        .SelectCommand.Parameters.Add("@TypeID1", SqlDbType.Int).Value = CInt(TypeID1)
    '    '        .Fill(dt)
    '    '    End With
    '    '    If Not da Is Nothing Then da.Dispose()
    '    '    If Not dt Is Nothing Then dt.Dispose()
    '    'Catch ex As Exception
    '    '    Throw ex
    '    'End Try

    '    If dt.Rows.Count > 0 Then
    '        dl_typeid2.Items.Clear()
    '        dl_typeid2.DataSource = dt
    '        dl_typeid2.DataValueField = "TypeID2"
    '        dl_typeid2.DataTextField = "TypeID2Name"
    '        dl_typeid2.DataBind()
    '        dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
    '        dl_typeid2.SelectedIndex = 0
    '        dl_typeid2.Enabled = True
    '    Else
    '        dl_typeid2.Items.Clear()
    '        dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
    '        dl_typeid2.SelectedIndex = 0
    '        dl_typeid2.Enabled = False
    '    End If
    'End Sub

    '代入資料
    Private Sub loadData(ByVal orgid As String)

        Dim dt As New DataTable
        Dim sql As String = ""
        sql &= " select a.orgid,a.orgname,a.comidno" & vbCrLf
        sql &= " ,a.orgkind1" & vbCrLf
        sql &= " ,a.orgzipcode" & vbCrLf
        sql &= " ,a.orgzipCODE6W" & vbCrLf
        sql &= " ,a.orgaddress" & vbCrLf
        sql &= " ,c.zipname" & vbCrLf
        sql &= " ,b.typeid1" & vbCrLf
        'sql &= " ,dbo.DECODE6(b.typeid1,'1','勞工在職進修計畫','2','勞工團體辦理勞工在職進修計畫','') plantype " & vbCrLf
        sql &= " ,v1.name plantype" & vbCrLf
        sql &= " ,b.typeid2" & vbCrLf
        sql &= " ,b.typeid2name" & vbCrLf
        sql &= " ,b.typeid2+'-'+b.typeid2name orgtype" & vbCrLf
        sql &= " ,d.ctname " & vbCrLf
        sql &= " FROM ORG_ORGINFO a" & vbCrLf
        sql &= " LEFT JOIN KEY_ORGTYPE1 b on b.ORGTYPEID1=a.ORGKIND1" & vbCrLf
        sql &= " left join v_OrgKind1 v1 on v1.sort= b.typeid1" & vbCrLf
        sql &= " left join id_zip c on c.zipcode=a.orgzipcode" & vbCrLf
        sql &= " left join id_city d on d.ctid=c.ctid" & vbCrLf
        sql &= " where orgid=@orgid" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("@orgid", SqlDbType.VarChar).Value = orgid
            dt.Load(.ExecuteReader())
        End With

        msg.Text = "查無資料!"

        tb_orgname.Text = ""
        tb_comidno.Text = ""
        dl_typeid1.SelectedIndex = 0
        dl_typeid2.SelectedIndex = 0

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = Nothing
            dr = dt.Rows(0)
            msg.Text = ""
            tb_orgname.Text = Convert.ToString(dr("orgname"))
            tb_comidno.Text = Convert.ToString(dr("comidno"))

            If Convert.ToString(dr("typeid1")) = "" Then
                dl_typeid1.SelectedIndex = 0
                dl_typeid2.Enabled = False
            Else
                Common.SetListItem(dl_typeid1, $"{dr("typeid1")}")
                TIMS.GET_DDL_TYPEID12(dl_typeid2, objconn, $"{dr("typeid1")}")
                If $"{dr("typeid2")}" <> "" Then
                    Common.SetListItem(dl_typeid2, $"{dr("typeid2")}")
                    'dl_typeid2.SelectedValue = $"{dr("typeid2")}"
                Else
                    dl_typeid2.SelectedIndex = 0
                End If
            End If
            city_code.Value = Convert.ToString(dr("orgzipcode"))
            hidZipCODE6W.Value = Convert.ToString(dr("orgzipCODE6W"))
            ZipCODEB3.Value = TIMS.GetZIPCODEB3(hidZipCODE6W.Value)
            TBCity.Text = TIMS.GET_FullCCTName(objconn, city_code.Value, hidZipCODE6W.Value)
            TBaddress.Text = Convert.ToString(dr("orgaddress"))

            '地址有資料
            Dim fgAddressHaveData As Boolean = (city_code.Value <> "" AndAlso ZipCODEB3.Value <> "" AndAlso TBaddress.Text <> "")
            If fgAddressHaveData AndAlso sm.UserInfo.LID = 1 Then
                dl_typeid1.Enabled = False
                dl_typeid2.Enabled = False
                TIMS.Display_None(city_zip)
                city_code.Disabled = True
                hidZipCODE6W.Disabled = True
                ZipCODEB3.Disabled = True
                TBCity.Enabled = False
                TBaddress.Enabled = False
                Dim tit1 As String = "不提供分署修改"
                TIMS.Tooltip(dl_typeid1, tit1)
                TIMS.Tooltip(dl_typeid2, tit1)
                TIMS.Tooltip(city_code, tit1)
                TIMS.Tooltip(hidZipCODE6W, tit1)
                TIMS.Tooltip(ZipCODEB3, tit1)
                TIMS.Tooltip(TBCity, tit1)
                TIMS.Tooltip(TBaddress, tit1)
            End If
        End If

    End Sub
#End Region

#Region "Function"
    '判斷是否為正整數
    Function IsInt(ByVal chkstr As String) As Boolean
        Dim bolRtn As Boolean = False
        Try
            If Int32.Parse(chkstr) = chkstr AndAlso Int32.Parse(chkstr) > 0 Then
                Return True
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return bolRtn
    End Function

#End Region

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

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            bt_save.Enabled = False
            Common.MessageBox(Me, "此功能目前僅開放給產業人才投資計劃與充電起飛計畫(在職)!!")
            Exit Sub
        End If
        'If Not TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    bt_save.Enabled = False
        '    Common.MessageBox(Me, "此功能目前僅開放給產業人才投資計劃!!")
        '    Exit Sub
        'End If

        If Not IsPostBack Then
            Call cCreate1()
        End If
    End Sub

    Sub cCreate1()
        dl_typeid1 = TIMS.Get_DDLPlanPoint0(Me, dl_typeid1, objconn)

        dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
        dl_typeid2.SelectedIndex = 0
        dl_typeid2.Enabled = False

        hid_orgid.Value = TIMS.ClearSQM(Me.Request("orgid"))
        loadData(hid_orgid.Value)

        'Litcity_code.Text = TIMS.Get_WorkZIPB3Link2()

        city_code.Attributes.Add("onblur", "getZipName('TBCity',this,this.value);")
        bt_save.Attributes.Add("onclick", "return chkSave();")

        Litcity_code.Text = TIMS.Get_WorkZIPB3Link2()

        Dim city_zip_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code, ZipCODEB3, hidZipCODE6W, TBCity, TBaddress)
        city_zip.Attributes.Add("onclick", city_zip_Attr_VAL)
    End Sub

    Private Sub dl_typeid1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dl_typeid1.SelectedIndexChanged
        Dim v_dl_typeid1 As String = TIMS.GetListValue(dl_typeid1)
        TIMS.GET_DDL_TYPEID12(dl_typeid2, objconn, v_dl_typeid1)
    End Sub

    Function Get_ORGTYPEID1() As String
        Dim v_dl_typeid1 As String = TIMS.GetListValue(dl_typeid1)
        Dim v_dl_typeid2 As String = TIMS.GetListValue(dl_typeid2)

        Dim sOrgTypeID1 As String = ""
        Dim sql As String = ""
        sql = "select orgtypeid1 from Key_OrgType1 where typeid1=@typeid1 and typeid2=@typeid2"
        'parms.Clear()
        Dim parms As New Hashtable From {{"typeid1", v_dl_typeid1}, {"typeid2", v_dl_typeid2}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count <> 1 Then
            Return sOrgTypeID1
            'Common.MessageBox(Me, "請選擇正確的計畫別 與 機構別!!") 'Exit Sub
        End If
        sOrgTypeID1 = dt.Rows(0)("orgtypeid1")
        Return sOrgTypeID1
        'If sOrgTypeID1 = "" Then'    Common.MessageBox(Me, "請選擇正確的計畫別 與 機構別!!")'    Exit Sub'End If   
    End Function

    Sub SaveData1(ByRef sOrgTypeID1 As String)
        hidZipCODE6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZipCODEB3.Value)

        Dim intCnt As Integer = 0
        Dim sql As String = ""
        sql &= " update org_orginfo"
        sql &= " set ModifyAcct=@ModifyAcct,ModifyDate=getdate() "
        sql &= " ,ORGKIND1=@orgkind1,orgzipcode=@orgzipcode"
        sql &= " ,orgzipCODE6W=@orgzipCODE6W"
        sql &= " ,orgaddress=@orgaddress "
        sql &= " where orgid=@orgid"
        'u_parms.Clear()
        Dim u_parms As New Hashtable From {
            {"ModifyAcct", sm.UserInfo.UserID},
            {"orgkind1", sOrgTypeID1},
            {"orgzipcode", city_code.Value},
            {"orgzipCODE6W", hidZipCODE6W.Value},
            {"orgaddress", TBaddress.Text},
            {"orgid", hid_orgid.Value}
        }
        intCnt = DbAccess.ExecuteNonQuery(sql, objconn, u_parms)

        'intCnt = 1
        'Try
        '    If Not dt Is Nothing Then dt.Dispose()
        '    If Not da Is Nothing Then da.Dispose()
        'Catch ex As Exception
        '    Throw ex
        'End Try

        If intCnt = 1 Then
            Me.Page.RegisterStartupScript("alertMsg", "<script language='javascript'>alert('儲存成功!');location.href='TC_01_017.aspx'</script>")
        End If
    End Sub

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            bt_save.Enabled = False
            Common.MessageBox(Me, "此功能目前僅開放給產業人才投資計劃與充電起飛計畫(在職)!!")
            Exit Sub
        End If
        'If Not TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    bt_save.Enabled = False
        '    Common.MessageBox(Me, "此功能目前僅開放給產業人才投資計劃!!")
        '    Exit Sub
        'End If

        Dim sOrgTypeID1 As String = Get_ORGTYPEID1()
        If sOrgTypeID1 = "" Then
            Common.MessageBox(Me, "請選擇正確的計畫別 與 機構別!!")
            Exit Sub
        End If

        If sm.UserInfo.LID <> 1 Then
            Dim errmsg1 As String = ""
            city_code.Value = TIMS.ClearSQM(city_code.Value)
            TBaddress.Text = TIMS.ClearSQM(TBaddress.Text)
            If (city_code.Value = "") Then errmsg1 &= "立案地址/會址-郵遞區號前3碼不可為空" & vbCrLf
            If (TBaddress.Text = "") Then errmsg1 &= "立案地址/會址-地址不可為空 " & vbCrLf
            If (city_code.Value <> "" AndAlso city_code.Value.Length < 3) Then errmsg1 &= String.Concat("立案地址/會址-郵遞區號前3碼-格式有誤!-", city_code.Value, vbCrLf)
            If (TBaddress.Text <> "" AndAlso TBaddress.Text.Length < 3) Then errmsg1 &= String.Concat("立案地址/會址-地址格式-內容有誤!-", TBaddress.Text, vbCrLf)
            If errmsg1 <> "" Then
                Common.MessageBox(Me, errmsg1)
                Exit Sub
            End If
        End If

        Call SaveData1(sOrgTypeID1)
    End Sub

    Private Sub bt_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_back.Click
        Me.Page.RegisterStartupScript("back", "<script language='javascript'>location.href='TC_01_017.aspx'</script>")
    End Sub
End Class
