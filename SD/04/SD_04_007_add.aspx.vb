Partial Class SD_04_007_add
    Inherits AuthBasePage

    'Dim objconn As SqlConnection
    'Dim sqlAdapter As SqlDataAdapter
    'Dim SqlStr As String
    'Dim dr As DataRow
    'Dim objtable As DataTable
    'Dim objstr As String
    ''----------------200810 Andy  start
    'Dim RID As String
    'Dim SetingYear As String
    '----------------          end

    'Class_RestTime
    '1.by 班級(by ocid)   依DistID.RID.OrgID.OCID
    '2.by 機構(分區 by rid) 依TPlanID.DistID.RID.OrgID 且ocid is null
    '3.by 機構(查詢無結果時) 依DistID.RID.OrgID

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'If saved.Value = "Y" Then
        '    But1.Enabled = False
        'Else
        '    But1.Enabled = True
        'End If
        'But1.Attributes("OnClick") = "if (check()==false){ return false;} else { this.disabled=true;document.getElementById('saved').value='Y';setTimeout(""document.getElementById('lnk_save').click()"",500); }"

        If Not IsPostBack Then
            'Hid_saved.Value = ""
            'But1.Attributes.Add("OnClick", "return check();")
            But1.Attributes.Add("onclick", "return xBtnSave1();")
            'But1.Attributes.Add("onclick", "xBtnSave1();")

            '----------------200810 Andy  start
            RIDValue.Value = Me.Request("RID").ToString  '單位Rid
            OrgID.Value = Me.Request("OrgID")
            OCID.Value = Me.Request("OCID")
            'SetingYear = Me.Request("Year").ToString      '設定年份
            PlanYear.Text = Me.Request("Year").ToString

            RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
            OrgID.Value = TIMS.ClearSQM(OrgID.Value)
            OCID.Value = TIMS.ClearSQM(OCID.Value)
            PlanYear.Text = TIMS.ClearSQM(PlanYear.Text)

            If OCID.Value.ToString <> "" Then             '班級
                TR_Class.Visible = True
            Else
                TR_Class.Visible = False
            End If

            LoadOrgClassName()
            '----------------    
            Me.ViewState("RestTime") = Session("RestTime")
            Session("RestTime") = Nothing

            Call LoadData()
        End If
    End Sub

    Function ChkRestTimeSet(ByVal sqlstr As String, Optional ByVal chktyp As String = "") As DataTable
        Dim rst As New DataTable

        TIMS.OpenDbConn(objConn)
        Dim sCmd As New SqlCommand(sqlstr, objConn)
        With sCmd
            .Parameters.Clear()
            Select Case chktyp
                Case "class"    'by 班級
                    .Parameters.Add("OCID", SqlDbType.Decimal).Value = OCID.Value
                    .Parameters.Add("DistID", SqlDbType.NVarChar).Value = Me.sm.UserInfo.DistID
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    .Parameters.Add("OrgID", SqlDbType.Decimal).Value = Me.OrgID.Value
                Case "NotSet"   '未設定時
                    .Parameters.Add("DistID", SqlDbType.NVarChar).Value = Me.sm.UserInfo.DistID
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    .Parameters.Add("OrgID", SqlDbType.Decimal).Value = Me.OrgID.Value
                Case "rid"      'by 機構
                    .Parameters.Add("DistID", SqlDbType.NVarChar).Value = Me.sm.UserInfo.DistID
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    .Parameters.Add("OrgID", SqlDbType.Decimal).Value = Me.OrgID.Value
            End Select

            rst.Load(.ExecuteReader())
        End With
        Return rst
    End Function

    Sub LoadData()
        '20080818 andy 以班為單位,20081029 改為-->若班未設定則以其機構設定為預設值(依計畫區分)
        '------------------------
        Dim dt As New DataTable
        Dim blnIsSeted As Boolean = False
        'Dim SqlNotSet, SqlClass, SqlRid, SqlOrg As String

        'by 班級(by ocid)
        Dim SqlClass As String = ""
        SqlClass = " select f.ocid,f.orgid, b.DistID,"
        SqlClass += " b.OrgLevel,b.Relship,b.rid,c.orgname "
        SqlClass += " from  Org_OrgPlanInfo  a " & vbCrLf
        SqlClass += " join  Auth_Relship b  on   a.RSID=b.RSID " & vbCrLf
        SqlClass += " join  Org_OrgInfo  c  on   c.OrgID=b.OrgID " & vbCrLf
        SqlClass += " join  Class_ClassInfo d  on  b.RID=d.RID " & vbCrLf
        SqlClass += " join  Class_RestTime  f  on  f.orgid=c.orgid " & vbCrLf
        SqlClass += " and   f.ocid=d.ocid " & vbCrLf
        SqlClass += " where 0=0 "
        '----------------------
        'SqlClass += " and f.TPlanID='" & Me.sm.UserInfo.TPlanID & "'"   '依計畫區分以登入後的 sm.UserInfo.TPlanID為依據
        '----------------------
        SqlClass += " and f.ocid=@OCID "
        SqlClass += " and b.DistID=@DistID "
        SqlClass += " and b.RID=@RID "
        SqlClass += " and f.OrgID=@OrgID "
        '------------------
        SqlClass += " group  by f.ocid,f.orgid, b.DistID,"
        SqlClass += " b.OrgLevel,b.Relship,b.rid,c.orgname "

        'by 機構(分區 by rid)  檢查是否Class_RestTime是否有該planid的該筆rid資料
        Dim SqlRid As String = ""
        SqlRid = "  select f.orgid, b.DistID,"
        SqlRid += " b.OrgLevel,b.Relship,b.rid,c.orgname "
        SqlRid += " from  Org_OrgPlanInfo a"
        SqlRid += " join  Auth_Relship b  on  a.RSID=b.RSID "
        SqlRid += " join  Org_OrgInfo  c  on  c.OrgID=b.OrgID "
        SqlRid += " join  Class_RestTime  f  on  f.orgid=c.orgid "
        SqlRid += " where 0=0 "
        '------------------
        SqlRid += " and f.TPlanID='" & Me.sm.UserInfo.TPlanID & "'"
        'SqlRid += " and f.Years='" & Right(PlanYear.Text, 2).ToString() & "'"
        '----------------------
        SqlRid += " and b.DistID=@DistID "
        SqlRid += " and f.RID=@RID "
        SqlRid += " and f.ocid is null"
        SqlRid += " and f.OrgID=@OrgID "
        '------------------
        SqlRid += " group  by f.ocid,f.orgid, b.DistID,"
        SqlRid += " b.OrgLevel,b.Relship,b.rid,c.orgname "

        If OCID.Value.ToString <> "" Then                             'by 班級  
            dt = ChkRestTimeSet(SqlClass, "class")
            Me.ViewState("SetingBy") = "NotSet"       '未設定時

            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Me.ViewState("SetingBy") = "class"
                Else
                    dt = ChkRestTimeSet(SqlRid, "rid")
                    Me.ViewState("SetingBy") = "NotSet"       '未設定時
                    If Not dt Is Nothing Then
                        If dt.Rows.Count > 0 Then
                            Me.ViewState("SetingBy") = "rid"         'by 機構(分區)
                        End If
                    End If
                End If
            End If
        Else
            dt = ChkRestTimeSet(SqlRid, "rid")
            Me.ViewState("SetingBy") = "NotSet"       '未設定時
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Me.ViewState("SetingBy") = "rid"          'by 機構(分區)
                End If
            End If
        End If

        '1.by 班級(by ocid)   依DistID.RID.OrgID.OCID (class)
        SqlClass = ""
        SqlClass &= " select f.ocid,f.orgid, b.DistID"
        SqlClass += " ,b.OrgLevel,b.Relship,b.rid,c.orgname "
        SqlClass += " , d.ClassCName + '(第' + d.CyclType + '期)' ClassCName " & vbCrLf
        SqlClass += " ,f.C1,f.C2,f.C3,f.C4,f.C5,f.C6,f.C7,f.C8,f.C9,f.C10,f.C11,f.C12  "
        SqlClass += " from Org_OrgPlanInfo  a " & vbCrLf

        SqlClass += " join  Auth_Relship b  on   a.RSID=b.RSID  " & vbCrLf
        SqlClass += " join  Org_OrgInfo  c  on   c.OrgID=b.OrgID " & vbCrLf
        SqlClass += " join  Class_ClassInfo d  on  b.RID=d.RID " & vbCrLf
        SqlClass += " join  Class_RestTime  f  on  f.orgid=c.orgid " & vbCrLf
        SqlClass += " and   f.ocid=d.ocid " & vbCrLf

        SqlClass += " where 0=0 "
        '------------------
        'SqlClass += " and f.TPlanID='" & Me.sm.UserInfo.TPlanID & "'"
        '------------------
        SqlClass += " and f.ocid=@OCID "
        SqlClass += " and b.DistID=@DistID "
        SqlClass += " and b.RID=@RID "
        SqlClass += " and f.OrgID=@OrgID "
        '------------------

        '2.by 機構(分區 by rid) 依TPlanID.DistID.RID.OrgID 且ocid is null (rid)
        SqlRid = "  select  f.ocid,f.orgid," & vbCrLf
        SqlRid += " b.OrgLevel,b.Relship,b.rid,c.orgname,c.orgid ," & vbCrLf
        SqlRid += " b.DistID, C.OrgID " & vbCrLf
        SqlRid += " ,f.C1,f.C2,f.C3,f.C4,f.C5,f.C6,f.C7,f.C8,f.C9,f.C10,f.C11,f.C12  "
        SqlRid += " from Org_OrgPlanInfo  a " & vbCrLf
        SqlRid += " join  Auth_Relship b  on   a.RSID=b.RSID  " & vbCrLf
        SqlRid += " join  Org_OrgInfo  c  on   c.OrgID=b.OrgID " & vbCrLf
        SqlRid += " join  Class_RestTime  f  on  f.orgid=c.orgid " & vbCrLf
        SqlRid += " where 0=0 "
        '------------------
        SqlRid += " and f.TPlanID='" & Me.sm.UserInfo.TPlanID & "'"
        '------------------
        SqlRid += " and f.ocid is null "
        SqlRid += " and b.DistID=@DistID "
        SqlRid += " and f.RID=@RID "
        SqlRid += " and f.OrgID=@OrgID "

        '3.by 機構(查詢無結果時) 依DistID.RID.OrgID (NotSet)
        Dim SqlNotSet As String = ""
        SqlNotSet = ""
        SqlNotSet += " select C.OrgID , B.DistID, B.OrgLevel, B.Relship, B.rid, C.orgname "
        SqlNotSet += " from  Org_OrgPlanInfo a"
        SqlNotSet += " join  Auth_Relship b  on  a.RSID=b.RSID "
        SqlNotSet += " join  Org_OrgInfo  c  on  c.OrgID=b.OrgID "
        SqlNotSet += " where 0 = 0 "
        SqlNotSet += " and b.DistID=@DistID"
        SqlNotSet += " and b.RID=@RID"
        SqlNotSet += " and c.OrgID=@OrgID "
        SqlNotSet += " group  by "
        SqlNotSet += " C.OrgID, B.DistID, B.OrgLevel, B.Relship, B.rid, C.orgname "

        Select Case Me.ViewState("SetingBy")
            Case "class"
                TD_Class.BgColor = "#ffffcc"     '黃
                dt = ChkRestTimeSet(SqlClass, "class")
            Case "rid"
                TD_org.BgColor = "#8ECEF1"       '藍
                dt = ChkRestTimeSet(SqlRid, "rid")
            Case "NotSet"
                dt = ChkRestTimeSet(SqlNotSet, "NotSet")
        End Select

        Dim dr As DataRow = Nothing
        If dt IsNot Nothing Then
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                blnIsSeted = True
            End If
        End If

        If Me.ViewState("SetingBy") = "NotSet" Then
            blnIsSeted = False
        End If

        If blnIsSeted = True Then   '該班or(機構)已有設定
            Me.OrgName.Text = dr("OrgName")
            Me.OrgID.Value = dr("OrgID")

            Me.DistID.Text = ""
            If Convert.ToString(dr("DistID")) <> "" Then
                Me.DistID.Text = TIMS.Get_DistName1(dr("DistID"))
            End If

            If Convert.IsDBNull(dr("C1")) Or Convert.ToString(dr("C1")) = "" Then
                Me.C11.Text = ""
                Me.C12.Text = ""
                Me.C13.Text = ""
                Me.C14.Text = ""
            Else
                Me.C11.Text = Convert.ToString(dr("C1")).Substring(0, 2)
                Me.C12.Text = Convert.ToString(dr("C1")).Substring(3, 2)
                Me.C13.Text = Convert.ToString(dr("C1")).Substring(6, 2)
                Me.C14.Text = Convert.ToString(dr("C1")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C2")) Or Convert.ToString(dr("C2")) = "" Then
                Me.C21.Text = ""
                Me.C22.Text = ""
                Me.C23.Text = ""
                Me.C24.Text = ""
            Else
                Me.C21.Text = Convert.ToString(dr("C2")).Substring(0, 2)
                Me.C22.Text = Convert.ToString(dr("C2")).Substring(3, 2)
                Me.C23.Text = Convert.ToString(dr("C2")).Substring(6, 2)
                Me.C24.Text = Convert.ToString(dr("C2")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C3")) Or Convert.ToString(dr("C3")) = "" Then
                Me.C31.Text = ""
                Me.C32.Text = ""
                Me.C33.Text = ""
                Me.C34.Text = ""
            Else
                Me.C31.Text = Convert.ToString(dr("C3")).Substring(0, 2)
                Me.C32.Text = Convert.ToString(dr("C3")).Substring(3, 2)
                Me.C33.Text = Convert.ToString(dr("C3")).Substring(6, 2)
                Me.C34.Text = Convert.ToString(dr("C3")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C4")) Or Convert.ToString(dr("C4")) = "" Then
                Me.C41.Text = ""
                Me.C42.Text = ""
                Me.C43.Text = ""
                Me.C44.Text = ""
            Else
                Me.C41.Text = Convert.ToString(dr("C4")).Substring(0, 2)
                Me.C42.Text = Convert.ToString(dr("C4")).Substring(3, 2)
                Me.C43.Text = Convert.ToString(dr("C4")).Substring(6, 2)
                Me.C44.Text = Convert.ToString(dr("C4")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C5")) Or Convert.ToString(dr("C5")) = "" Then
                Me.C51.Text = ""
                Me.C52.Text = ""
                Me.C53.Text = ""
                Me.C54.Text = ""
            Else
                Me.C51.Text = Convert.ToString(dr("C5")).Substring(0, 2)
                Me.C52.Text = Convert.ToString(dr("C5")).Substring(3, 2)
                Me.C53.Text = Convert.ToString(dr("C5")).Substring(6, 2)
                Me.C54.Text = Convert.ToString(dr("C5")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C6")) Or Convert.ToString(dr("C6")) = "" Then
                Me.C61.Text = ""
                Me.C62.Text = ""
                Me.C63.Text = ""
                Me.C64.Text = ""
            Else
                Me.C61.Text = Convert.ToString(dr("C6")).Substring(0, 2)
                Me.C62.Text = Convert.ToString(dr("C6")).Substring(3, 2)
                Me.C63.Text = Convert.ToString(dr("C6")).Substring(6, 2)
                Me.C64.Text = Convert.ToString(dr("C6")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C7")) Or Convert.ToString(dr("C7")) = "" Then
                Me.C71.Text = ""
                Me.C72.Text = ""
                Me.C73.Text = ""
                Me.C74.Text = ""
            Else
                Me.C71.Text = Convert.ToString(dr("C7")).Substring(0, 2)
                Me.C72.Text = Convert.ToString(dr("C7")).Substring(3, 2)
                Me.C73.Text = Convert.ToString(dr("C7")).Substring(6, 2)
                Me.C74.Text = Convert.ToString(dr("C7")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C8")) Or Convert.ToString(dr("C8")) = "" Then
                Me.C81.Text = ""
                Me.C82.Text = ""
                Me.C83.Text = ""
                Me.C84.Text = ""
            Else
                Me.C81.Text = Convert.ToString(dr("C8")).Substring(0, 2)
                Me.C82.Text = Convert.ToString(dr("C8")).Substring(3, 2)
                Me.C83.Text = Convert.ToString(dr("C8")).Substring(6, 2)
                Me.C84.Text = Convert.ToString(dr("C8")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C9")) Or Convert.ToString(dr("C9")) = "" Then
                Me.C91.Text = ""
                Me.C92.Text = ""
                Me.C93.Text = ""
                Me.C94.Text = ""
            Else
                Me.C91.Text = Convert.ToString(dr("C9")).Substring(0, 2)
                Me.C92.Text = Convert.ToString(dr("C9")).Substring(3, 2)
                Me.C93.Text = Convert.ToString(dr("C9")).Substring(6, 2)
                Me.C94.Text = Convert.ToString(dr("C9")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C10")) Or Convert.ToString(dr("C10")) = "" Then
                Me.C101.Text = ""
                Me.C102.Text = ""
                Me.C103.Text = ""
                Me.C104.Text = ""
            Else
                Me.C101.Text = Convert.ToString(dr("C10")).Substring(0, 2)
                Me.C102.Text = Convert.ToString(dr("C10")).Substring(3, 2)
                Me.C103.Text = Convert.ToString(dr("C10")).Substring(6, 2)
                Me.C104.Text = Convert.ToString(dr("C10")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C11")) Or Convert.ToString(dr("C11")) = "" Then
                Me.C111.Text = ""
                Me.C112.Text = ""
                Me.C113.Text = ""
                Me.C114.Text = ""
            Else
                Me.C111.Text = Convert.ToString(dr("C11")).Substring(0, 2)
                Me.C112.Text = Convert.ToString(dr("C11")).Substring(3, 2)
                Me.C113.Text = Convert.ToString(dr("C11")).Substring(6, 2)
                Me.C114.Text = Convert.ToString(dr("C11")).Substring(9, 2)
            End If
            If Convert.IsDBNull(dr("C12")) Or Convert.ToString(dr("C12")) = "" Then
                Me.C121.Text = ""
                Me.C122.Text = ""
                Me.C123.Text = ""
                Me.C124.Text = ""
            Else
                Me.C121.Text = Convert.ToString(dr("C12")).Substring(0, 2)
                Me.C122.Text = Convert.ToString(dr("C12")).Substring(3, 2)
                Me.C123.Text = Convert.ToString(dr("C12")).Substring(6, 2)
                Me.C124.Text = Convert.ToString(dr("C12")).Substring(9, 2)
            End If
        Else
            Me.C11.Text = ""
            Me.C12.Text = ""
            Me.C13.Text = ""
            Me.C14.Text = ""
            Me.C21.Text = ""
            Me.C22.Text = ""
            Me.C23.Text = ""
            Me.C24.Text = ""
            Me.C31.Text = ""
            Me.C32.Text = ""
            Me.C33.Text = ""
            Me.C34.Text = ""
            Me.C41.Text = ""
            Me.C42.Text = ""
            Me.C43.Text = ""
            Me.C44.Text = ""
            Me.C51.Text = ""
            Me.C52.Text = ""
            Me.C53.Text = ""
            Me.C54.Text = ""
            Me.C61.Text = ""
            Me.C62.Text = ""
            Me.C63.Text = ""
            Me.C64.Text = ""
            Me.C71.Text = ""
            Me.C72.Text = ""
            Me.C73.Text = ""
            Me.C74.Text = ""
            Me.C81.Text = ""
            Me.C82.Text = ""
            Me.C83.Text = ""
            Me.C84.Text = ""
            Me.C91.Text = ""
            Me.C92.Text = ""
            Me.C93.Text = ""
            Me.C94.Text = ""
            Me.C101.Text = ""
            Me.C102.Text = ""
            Me.C103.Text = ""
            Me.C104.Text = ""
            Me.C111.Text = ""
            Me.C112.Text = ""
            Me.C113.Text = ""
            Me.C114.Text = ""
            Me.C121.Text = ""
            Me.C122.Text = ""
            Me.C123.Text = ""
            Me.C124.Text = ""
        End If

    End Sub

    Sub LoadOrgClassName()
        'Dim da As New SqlDataAdapter
        'Dim ds As New DataSet
        'Dim dt As New DataTable
        'Dim dr As DataRow
        'Me.ViewState("IsLoaded") = "Y"
        Dim sqlstr As String = ""
        If OCID.Value.ToString <> "" Then    '班級
            sqlstr = ""
            sqlstr &= " select a.rid, a.ocid, f.DistID, e.OrgName,e.OrgID"
            sqlstr &= " ,a.ClassCName + '(第' + a.CyclType + '期)' ClassCName"
            sqlstr &= " ,ar2.OrgName2"
            sqlstr &= " FROM Class_ClassInfo a "
            sqlstr &= " join Auth_Relship f  on a.RID=f.RID "
            sqlstr &= " join Org_OrgInfo e on f.OrgID=e.Orgid "
            sqlstr &= " LEFT JOIN MVIEW_RELSHIP23 ar2 on ar2.RID3=a.RID " & vbCrLf
            sqlstr &= " where 1=1 "
            sqlstr &= " and a.ocid =@OCID "
            sqlstr &= " and f.rid=@RID "

        Else
            sqlstr = ""
            sqlstr &= " select a.DistID,b.OrgID,b.Orgname"
            sqlstr &= " from Auth_Relship a"
            sqlstr &= " join Org_orginfo b on a.orgid=b.orgid"
            sqlstr &= " where a.RID=@RID "
        End If

        Dim sCmd As New SqlCommand(sqlstr, objConn)
        TIMS.OpenDbConn(objConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            If OCID.Value.ToString <> "" Then   '班級
                .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID.Value
            End If
            .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
            dt.Load(.ExecuteReader())
        End With


        'Try
        '    With da
        '        da.SelectCommand = New SqlCommand(sqlstr, objConn)
        '        If OCID.Value.ToString <> "" Then   '班級
        '            da.SelectCommand.Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID.Value.ToString
        '            da.SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value.ToString
        '        Else
        '            da.SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value.ToString
        '        End If
        '        .Fill(ds, "QueryTB")
        '        dt = ds.Tables("QueryTB")
        '    End With
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString())
        'End Try

        '20081027 Andy
        '----------------
        TR_CtrlOrg.Visible = False

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            If OCID.Value <> "" Then   '班級
                ClassName.Text = dr("ClassCName")
                If Convert.ToString(dr("OrgName2")) <> "" Then
                    TR_CtrlOrg.Visible = True
                    CtrlOrg.Text = dr("OrgName2")
                End If
            End If

            OrgID.Value = dr("OrgID")
            OrgName.Text = dr("OrgName")
            Me.DistID.Text = ""
            If Convert.ToString(dr("DistID")) <> "" Then
                Me.DistID.Text = TIMS.Get_DistName1(dr("DistID"))
            End If
        End If

        'Dim RelshipTable As DataTable
        'sqlstr = ""
        'sqlstr &= " SELECT a.RID,a.Relship,b.OrgName "
        'sqlstr &= " FROM Auth_Relship a "
        'sqlstr &= " JOIN Org_OrgInfo b ON a.OrgID=b.OrgID "

        'da.SelectCommand = New SqlCommand(sqlstr, objConn)
        'da.Fill(ds, "CtrOrgName")
        'RelshipTable = ds.Tables("CtrOrgName")

        'If RIDValue.Value.ToString.Length <> 1 Then
        '    If RelshipTable.Select("RID='" & RIDValue.Value.ToString & "'").Length <> 0 Then
        '        Dim Relship As String
        '        Dim Parent As String
        '        Relship = RelshipTable.Select("RID='" & RIDValue.Value.ToString & "'")(0)("Relship")
        '        Parent = Split(Relship, "/")(Split(Relship, "/").Length - 3)
        '        If RelshipTable.Select("RID='" & Parent & "'").Length <> 0 Then
        '            CtrlOrg.Text = RelshipTable.Select("RID='" & Parent & "'")(0)("OrgName")
        '        End If
        '    End If
        'Else
        '    TR_CtrlOrg.Visible = False
        'End If
    End Sub

    'Private Sub But1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But1.Click
    '    But1.Enabled = False
    'End Sub

    '回上一頁
    Protected Sub But2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But2.Click
        Session("RestTime") = Me.ViewState("RestTime")
        Me.ViewState("IsLoaded") = "Y"
        TIMS.Utl_Redirect1(Me, "SD_04_007.aspx?ID=" & Request("ID"))
    End Sub

    Public Shared Function GetTIMECOMBINE(ByRef sHH1 As String, ByRef sMM1 As String, ByRef sHH2 As String, ByRef sMM2 As String) As Object
        Dim rst As Object = System.DBNull.Value 'Convert.DBNull()
        If String.IsNullOrEmpty(sHH1) Then Return rst
        If String.IsNullOrEmpty(sMM1) Then Return rst
        If String.IsNullOrEmpty(sHH2) Then Return rst
        If String.IsNullOrEmpty(sMM2) Then Return rst
        rst = String.Concat(sHH1, ":", sMM1, "~", sHH2, ":", sMM2)
        Return rst
    End Function

    Sub SaveData1()
        'Dim sqlAdapter As SqlDataAdapter
        'Dim sqlstr_update As String
        'Dim sqlTable As New DataTable
        'Dim atable As New DataTable
        Dim sqlstr As String = ""
        sqlstr = "Select * from Class_RestTime Where OrgID=" & Me.OrgID.Value & " and RID='" & RIDValue.Value.ToString() & "' and  TPlanID='" & Me.sm.UserInfo.TPlanID & "' and OCID is null  "
        If OCID.Value.ToString <> "" Then          'by 班級  
            Select Case Me.ViewState("SetingBy")   '設定值(若未設定)依層級往上來找  班級-->機構(分區) 
                Case "class"
                    sqlstr = "Select * from Class_RestTime Where OrgID=" & Me.OrgID.Value & " and OCID=" & OCID.Value.ToString & ""
                Case "rid"
                    sqlstr = "Select * from Class_RestTime Where OrgID=" & Me.OrgID.Value & " and RID='" & RIDValue.Value.ToString() & "' and  TPlanID='" & Me.sm.UserInfo.TPlanID & "'  and OCID is null  "
                Case Else
                    sqlstr = "Select * from Class_RestTime Where OrgID=" & Me.OrgID.Value & " and OCID=" & OCID.Value.ToString & ""
            End Select
        End If

        '20080818 Andy 改以班別為單位
        Dim dr As DataRow = DbAccess.GetOneRow(sqlstr, objConn)

        Dim sqlstrdel As String = ""
        If OCID.Value.ToString <> "" Then
            'by 班級  
            If Me.ViewState("SetingBy") = "class" Then
                If dr IsNot Nothing Then
                    sqlstrdel = "delete Class_RestTime where CRTID=" & dr("CRTID")
                    DbAccess.ExecuteNonQuery(sqlstrdel, objConn)
                End If
            End If
        Else
            'by 機構
            Select Case Me.ViewState("SetingBy")
                Case "rid", "org"
                    sqlstrdel = "delete Class_RestTime where CRTID=" & dr("CRTID")
                    DbAccess.ExecuteNonQuery(sqlstrdel, objConn)
            End Select
        End If

        '清除所有可能重複的資料
        If OCID.Value.ToString <> "" Then   'by 班級  
            sqlstrdel = "DELETE Class_RestTime Where 1=1 AND OCID='" & OCID.Value.ToString & "' "
            DbAccess.ExecuteNonQuery(sqlstrdel, objConn)
        Else
            If OrgID.Value <> "" AndAlso Me.RIDValue.Value <> "" Then
                sqlstrdel = ""
                sqlstrdel &= " DELETE Class_RestTime"
                sqlstrdel &= " Where 1=1 "
                sqlstrdel &= " AND OrgID='" & OrgID.Value & "'"
                sqlstrdel &= " AND RID='" & Me.RIDValue.Value & "'"
                DbAccess.ExecuteNonQuery(sqlstrdel, objConn)
            End If
        End If

        Dim da As SqlDataAdapter = Nothing
        Dim sqlTable As DataTable = Nothing
        sqlTable = DbAccess.GetDataTable("SELECT * FROM Class_RestTime WHERE 1<>1", da, objConn)
        dr = sqlTable.NewRow
        sqlTable.Rows.Add(dr)
        'CLASS_RESTTIME_CRTID_SEQ
        dr("CRTID") = DbAccess.GetNewId(objConn, "CLASS_RESTTIME_CRTID_SEQ,CLASS_RESTTIME,CRTID")
        If OCID.Value.ToString <> "" Then   'by 班級  
            dr("OrgID") = Convert.ToInt64(OrgID.Value.ToString())
            dr("RID") = Me.RIDValue.Value.ToString()
            dr("OCID") = Convert.ToInt32(OCID.Value.ToString)
        Else
            dr("OrgID") = Convert.ToInt64(OrgID.Value.ToString())
            dr("RID") = Me.RIDValue.Value.ToString()
            dr("OCID") = Convert.DBNull
        End If

        dr("TPlanID") = Convert.ToString(Me.sm.UserInfo.TPlanID)
        dr("RID") = RIDValue.Value.ToString()

        dr("C1") = GetTIMECOMBINE(C11.Text, C12.Text, C13.Text, C14.Text)
        dr("C2") = GetTIMECOMBINE(C21.Text, C22.Text, C23.Text, C24.Text)
        dr("C3") = GetTIMECOMBINE(C31.Text, C32.Text, C33.Text, C34.Text)
        dr("C4") = GetTIMECOMBINE(C41.Text, C42.Text, C43.Text, C44.Text)

        dr("C5") = GetTIMECOMBINE(C51.Text, C52.Text, C53.Text, C54.Text)
        dr("C6") = GetTIMECOMBINE(C61.Text, C62.Text, C63.Text, C64.Text)
        dr("C7") = GetTIMECOMBINE(C71.Text, C72.Text, C73.Text, C74.Text)
        dr("C8") = GetTIMECOMBINE(C81.Text, C82.Text, C83.Text, C84.Text)

        dr("C9") = GetTIMECOMBINE(C91.Text, C92.Text, C93.Text, C94.Text)
        dr("C10") = GetTIMECOMBINE(C101.Text, C102.Text, C103.Text, C104.Text)
        dr("C11") = GetTIMECOMBINE(C111.Text, C112.Text, C113.Text, C114.Text)
        dr("C12") = GetTIMECOMBINE(C121.Text, C122.Text, C123.Text, C124.Text)

        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(sqlTable, da)

        But1.Enabled = True

        Session("RestTime") = Me.ViewState("RestTime")
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "alert('作息時間設定-成功!!');" + vbCrLf
        strScript += "location.href='SD_04_007.aspx?ID=" & Request("ID") & "';" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    ''儲存(隱藏式儲存)
    'Protected Sub lnk_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lnk_save.Click
    '    Call SaveData1()
    'End Sub
    'Protected Sub But1_Click(sender As Object, e As EventArgs) Handles But1.Click
    '    Call lnk_save_Click()
    'End Sub

    '儲存
    Private Sub lnk_save_Click(sender As Object, e As System.EventArgs) Handles lnk_save.Click
        If Hid_saved.Value = "Y" Then
            Hid_saved.Value = "S"
            Call SaveData1() '儲存
            Hid_saved.Value = ""
            But1.Enabled = True
        End If
    End Sub


End Class
