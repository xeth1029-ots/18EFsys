Partial Class SD_08_014
    Inherits AuthBasePage

    Dim sql As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cache.SetExpires(DateTime.Now())
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        If Page.IsPostBack = False Then

            If Request("ID") Is Nothing Then
                ViewState("ID") = ""
            Else
                ViewState("ID") = Request("ID")
            End If


            Dim SUBID As String

            If Request("SUBID") Is Nothing Then
                SUBID = ""
            Else
                SUBID = Request("SUBID").Trim.Replace("'", "''")
            End If

            btnExit.Attributes("onclick") = "window.close();"

            '申請類別
            sql = "select IdentityID,name as Idname from Key_Identity where identityid in ('03','04','05','06','07','10','18')"
            TIMS.BindDDL(EIdentityID, sql, "IdentityID", "Idname", objconn)

            '申請障別
            sql = "select HandTypeID,name as Handname from Key_HandicatType"
            TIMS.BindDDL(EHandicat, sql, "HandTypeID", "Handname", objconn)

            '殘障等級
            sql = "select HandLevelID,name as HandTypename from Key_HandicatLevel"
            TIMS.BindDDL(EHandicatlevel, sql, "HandLevelID", "HandTypename", objconn)

            '參與職類
            sql = "select stid,jobname from Key_SubTrainType"
            TIMS.BindDDL(ETrainType, sql, "stid", "jobname", objconn)

            Call LoadData(SUBID)

        End If

    End Sub



    '載入資料
    Private Sub LoadData(ByVal SUBID As String)

        Dim dr As DataRow
        Dim dr1 As DataRow
        Dim sql As String

        ''sql = "select * from Sub_SubSidyApply_All where subid=" & SUBID
        'sql = "select a.*, b.orgname from Sub_SubSidyApply_All a left join sub_org b "
        'sql += "on a.OrgId=b.orgid where subid=" & SUBID
        sql = " select a.*, b.orgname,c.Reason "
        sql += " from Sub_SubSidyApply_All a left join sub_org b  on a.OrgId=b.orgid"
        sql += "                             left join Key_RejectTReason c on a.RTReason=c.RTReasonID"
        sql += " where SUBID = " & SUBID & " "
        dr = DbAccess.GetOneRow(sql, objconn)

        If Not dr Is Nothing Then

            '所屬單位
            If Convert.ToString(dr("fromtype")) = "1" Then
                sql = "select orgname from Org_OrgInfo where OrgId = '" & Convert.ToString(dr("OrgId")) & "'"
                dr1 = DbAccess.GetOneRow(sql, objconn)

                If Not dr1 Is Nothing Then
                    EOrgName.Text = Convert.ToString(dr1("orgname"))
                End If
            Else
                EOrgName.Text = Convert.ToString(dr("orgname"))
            End If
            EOrgId.Value = Convert.ToString(dr("orgid"))


            '身分證號
            EIdno.Text = dr("idno")

            '姓名
            EName.Text = Convert.ToString(dr("name"))

            '生日
            If Not Convert.IsDBNull(dr("birthday")) Then
                EBirthday.Text = Common.FormatDate2Roc(Convert.ToDateTime(dr("birthday")).ToString("yyyy/MM/dd"))
            End If

            '性別
            Common.SetListItem(ESex, Convert.ToString(dr("sex")))

            '通訊地址
            If Not Convert.IsDBNull(dr("ZipCode1")) Then
                sql = "select ic.CTName,iz.ZipName from ID_ZIP iz JOIN ID_City ic ON ic.CTID = iz.CTID where iz.ZipCode = '" & dr("ZipCode1") & "'"
                dr1 = DbAccess.GetOneRow(sql, objconn)

                If Not dr1 Is Nothing Then
                    City1.Text = "(" & dr("ZipCode1") & ")" & dr1("CTName") & dr1("ZipName")
                End If
            End If
            Address.Text = Convert.ToString(dr("Address"))


            '戶籍地址
            If Not Convert.IsDBNull(dr("ZipCode2")) Then
                sql = "select ic.CTName,iz.ZipName from ID_ZIP iz JOIN ID_City ic ON ic.CTID = iz.CTID where iz.ZipCode = '" & dr("ZipCode2") & "'"
                dr1 = DbAccess.GetOneRow(sql, objconn)

                If Not dr1 Is Nothing Then
                    City2.Text = "(" & dr("ZipCode2") & ")" & dr1("CTName") & dr1("ZipName")
                End If
            End If
            HouseholdAddress.Text = Convert.ToString(dr("HouseholdAddress"))


            '連絡電話
            EPhone.Text = dr("phone").ToString

            '申請類別
            Common.SetListItem(EIdentityID, Convert.ToString(dr("identityid")))

            '申請障別
            Common.SetListItem(EHandicat, Convert.ToString(dr("handicat")))

            '殘障等級
            Common.SetListItem(EHandicatlevel, Convert.ToString(dr("handicatlevel")))

            '審核狀態
            If Not Convert.IsDBNull(dr("AppliedStatusF")) Then
                If dr("AppliedStatusF") = "Y" Then
                    Me.lblAppliedStatusF.Text = "通過"
                Else
                    Me.lblAppliedStatusF.Text = "未通過"
                End If
            Else
                Me.lblAppliedStatusF.Text = "未審核"
            End If

            If Not Convert.IsDBNull(dr("isDownload")) Then
                If dr("isDownload") = "1" Then
                    Me.lblisDownload.Text = "已送"
                End If
            End If

            Select Case dr("AppliedStatusFin").ToString
                Case "Y"
                    lblAppliedStatusFin.Text = "通過"
                Case ""
                    lblAppliedStatusFin.Text = ""
                Case Else
                    lblAppliedStatusFin.Text = dr("AppliedStatusFin").ToString + "，" + Convert.ToString(dr("FailReasonFin"))
            End Select


            '訓練起迄
            If Not Convert.IsDBNull(dr("tsdate")) Then
                ETSDate.Text = Common.FormatDate2Roc(Convert.ToDateTime(dr("tsdate")).ToString("yyyy/MM/dd"))
            End If

            If Not Convert.IsDBNull(dr("tedate")) Then
                ETEDate.Text = Common.FormatDate2Roc(Convert.ToDateTime(dr("tedate")).ToString("yyyy/MM/dd"))
            End If

            If Not Convert.IsDBNull(dr("csdate")) Then
                CSDate.Text = Common.FormatDate2Roc(Convert.ToDateTime(dr("csdate")).ToString("yyyy/MM/dd"))
            End If


            '申請日期
            If Not Convert.IsDBNull(dr("applydate")) Then
                EApplydate.Text = Common.FormatDate2Roc(Convert.ToDateTime(dr("applydate")).ToString("yyyy/MM/dd"))
            End If

            '參訓職類
            Common.SetListItem(ETrainType, Convert.ToString(dr("traincode")))

            '參訓班別
            EClassName.Text = Convert.ToString(dr("classname"))

            '補助月數
            ETMonth.Text = Convert.ToString(dr("trainingmonth"))

            '申請金額
            ETMoney.Text = Convert.ToString(dr("trainingmoney"))

            '核發月數
            EAMonth.Text = Convert.ToString(dr("ApplyMonth"))

            '核發金額
            EAMoney.Text = Convert.ToString(dr("ApplyMoney"))

            '實際請領月數
            EPayMonth.Text = Convert.ToString(dr("PayMonth"))

            '實際請領金額
            EPayMoney.Text = Convert.ToString(dr("PayMoney"))


            '離退訓日期
            If Not Convert.IsDBNull(dr("LDate")) Then
                ELdate.Text = Common.FormatDate2Roc(Convert.ToDateTime(dr("LDate")).ToString("yyyy/MM/dd"))
            End If

            '離退訓
            Select Case Convert.ToString(dr("LFlag"))
                Case "1"
                    ELflag1.Checked = True
                Case "2"
                    ELflag2.Checked = True
            End Select


            '繳回金額
            ERtnMoney.Text = Convert.ToString(dr("RtnMoney"))

            'stella add 2006/12/15
            '職退原因
            RTReason.Text = Convert.ToString(dr("Reason"))

            '繳回月數
            RtnMonth.Text = Convert.ToString(dr("RtnMonth"))

            '離退審核結果
            If Not Convert.IsDBNull(dr("LVerify")) Then
                If dr("LVerify") = "Y" Then
                    LVerify.Text = "通過"
                Else
                    LVerify.Text = "不通過"
                End If
            Else
                If Not Convert.IsDBNull(dr("LDate")) Then
                    LVerify.Text = "未審核"
                End If
            End If
            'stella add 2006/12/15

        End If

    End Sub

End Class
