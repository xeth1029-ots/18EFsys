Partial Class OB_01_003_add
    Inherits AuthBasePage
#Region "FUNC1"

    Function CheckData(ByRef Errmag As String) As Boolean
        Errmag = ""
        CheckData = False

        If Me.ddl_years.SelectedValue = "" Then
            Errmag += "請選擇年度" & vbCrLf
        End If
        If Me.ddlTenderCName.SelectedValue = "" Then
            Errmag += "請選擇標案名稱" & vbCrLf
        End If

        MTSubject.Text = MTSubject.Text.Trim
        If MTSubject.Text = "" Then
            Errmag += "請輸入會議主題" & vbCrLf
        End If

        MTDate.Text = MTDate.Text.Trim
        If MTDate.Text = "" Then
            Errmag += "請輸入會議日期" & vbCrLf
        Else
            If Not IsDate(MTDate.Text) Then
                Errmag += "會議日期必須是西元年月日格式(yyyy/mm/dd) " & vbCrLf
            End If
        End If

        MTPlace.Text = MTPlace.Text.Trim
        If MTPlace.Text = "" Then
            Errmag += "請輸入會議地點" & vbCrLf
        End If

        MTContent.Text = MTContent.Text.Trim
        If MTContent.Text = "" Then
            Errmag += "請輸入會議議程內容" & vbCrLf
        End If

        Select Case ViewState("Action")
            Case "ADD"
                If Chk_EXISTS(MTSubject.Text) Then
                    Errmag += "此會議主題已經存在，請另設新會議主題" & vbCrLf
                End If
            Case "EDIT"
                If Chk_EXISTS(MTSubject.Text, ViewState("mtsn")) Then
                    Errmag += "此會議主題已經存在、請另設新會議主題" & vbCrLf
                End If
        End Select

        If Errmag = "" Then
            CheckData = True
        End If

    End Function

    Function Chk_EXISTS(ByVal MTSubject As String, Optional ByVal mtsn As Integer = 0) As Boolean
        Dim str_flag As String = Nothing
        Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql += " SELECT  " & vbCrLf
        'sql += " a.MTSN,a.MTSUBJECT,a.MTDATE,a.MTPLACE,dbms_lob.substr( a.MTCONTENT, 4000, 1 ) MTCONTENT,  " & vbCrLf
        'sql += " a.CREATEACCT,a.CREATETIME,a.MODIFYACCT,a.MODIFYTIME,a.DISTID,a.TSN  " & vbCrLf
        sql = "" & vbCrLf
        sql += " SELECT a.MTSN,a.MTSUBJECT,a.MTDATE,a.MTPLACE, a.MTCONTENT,  " & vbCrLf
        sql += " a.CREATEACCT,a.CREATETIME,a.MODIFYACCT,a.MODIFYTIME,a.DISTID,a.TSN  " & vbCrLf
        sql += " FROM OB_Meeting a " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If
        sql += " AND a.MTSubject='" & MTSubject & "'" & vbCrLf

        If mtsn <> 0 Then
            '排除本身
            sql += " AND a.mtsn !='" & mtsn & "' " & vbCrLf
        End If

        str_flag = DbAccess.ExecuteScalar(sql, objconn)
        If str_flag Is Nothing Then
            Chk_EXISTS = False
        Else
            Chk_EXISTS = True
        End If
    End Function
#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            ddl_years = TIMS.GetSyear(ddl_years, Year(Now) - 1, Year(Now) + 3, True)

            ViewState("Action") = TIMS.ClearSQM(UCase(Request("Action")))

            Dim sql As String = ""
            Dim dr As DataRow = Nothing
            ViewState("mtsn") = TIMS.ClearSQM(Request("mtsn"))

            If ViewState("mtsn") <> "" And Me.ViewState("Action") = "EDIT" Then
                sql = ""
                sql += " SELECT a.MTSN,a.MTSUBJECT,a.MTDATE,a.MTPLACE, a.MTCONTENT " & vbCrLf
                sql += " ,a.CREATEACCT,a.CREATETIME,a.MODIFYACCT,a.MODIFYTIME,a.DISTID,a.TSN  " & vbCrLf
                sql += " ,ot.TenderCName, ot.years" & vbCrLf
                sql += " FROM OB_Meeting a " & vbCrLf
                sql += " JOIN OB_tender ot on ot.tsn=a.tsn " & vbCrLf
                sql += " WHERE a.mtsn='" & ViewState("mtsn") & "' " & vbCrLf
                If sm.UserInfo.DistID <> "000" Then
                    sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
                End If
                dr = DbAccess.GetOneRow(sql, objconn)

                Me.ViewState("mtsn") = ViewState("mtsn")

                Common.SetListItem(ddl_years, Convert.ToString(dr("years")))
                TIMS.Get_TenderCName(ddlTenderCName, Convert.ToString(dr("years")), objconn)

                Common.SetListItem(ddlTenderCName, dr("tsn"))
                Me.MTSubject.Text = dr("MTSubject")
                Me.MTDate.Text = FormatDateTime(dr("MTDate"), DateFormat.ShortDate)
                Me.MTPlace.Text = dr("MTPlace")
                Me.MTContent.Text = dr("MTContent")

                ddl_years.Enabled = False
                ddlTenderCName.Enabled = False
            End If

        End If

        'PageLoadSetLast1()

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim Errmsg As String = ""
        If CheckData(Errmsg) Then
            Select Case Me.ViewState("Action")
                Case "ADD"
                    SAVE_Meeting()
                Case "EDIT"
                    SAVE_Meeting(Me.ViewState("mtsn"))
            End Select

        Else
            Common.MessageBox(Me, Errmsg)
        End If
    End Sub

    Sub SAVE_Meeting(Optional ByVal mtsn As Integer = 0)
        Dim OkErrorflag As Boolean = True

        Dim sqlCmd As SqlCommand
        'Dim objconn As SqlConnection
        Dim sqlStr As String = ""
        Select Case Me.ViewState("Action")
            Case "ADD"
                sqlStr = "" & vbCrLf
                sqlStr += " INSERT INTO OB_Meeting( tsn, MTSubject, MTDate, MTPlace, MTContent " & vbCrLf
                sqlStr += " , DistID, CreateAcct, CreateTime, ModifyAcct, ModifyTime) " & vbCrLf
                sqlStr += " VALUES( @tsn, @MTSubject, @MTDate , @MTPlace , @MTContent " & vbCrLf
                sqlStr += " , @DistID, @CreateAcct, getdate(), @ModifyAcct, getdate()) " & vbCrLf

            Case "EDIT"
                sqlStr = "" & vbCrLf
                sqlStr += " UPDATE  OB_Meeting" & vbCrLf
                sqlStr += " SET MTSubject=@MTSubject " & vbCrLf
                sqlStr += " , MTDate=@MTDate " & vbCrLf
                sqlStr += " , MTPlace=@MTPlace " & vbCrLf
                sqlStr += " , MTContent=@MTContent " & vbCrLf
                sqlStr += " , ModifyAcct=@ModifyAcct" & vbCrLf
                sqlStr += " , ModifyTime=getdate()" & vbCrLf
                sqlStr += " WHERE mtsn=@mtsn " & vbCrLf

        End Select

        Try
            'objconn = DbAccess.GetConnection
            'objconn.Open()

            TIMS.OpenDbConn(objconn)
            sqlCmd = New SqlCommand(sqlStr, objconn)
            With sqlCmd
                .Parameters.Clear()

                .Parameters.Add("MTSubject", SqlDbType.NVarChar, 20).Value = MTSubject.Text
                .Parameters.Add("MTDate", SqlDbType.DateTime).Value = MTDate.Text
                .Parameters.Add("MTPlace", SqlDbType.NVarChar, 20).Value = MTPlace.Text
                .Parameters.Add("MTContent", SqlDbType.NText).Value = MTContent.Text
                .Parameters.Add("ModifyAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID

                If Me.ViewState("Action") = "ADD" Then
                    .Parameters.Add("tsn", SqlDbType.Int).Value = Me.ddlTenderCName.SelectedValue
                    .Parameters.Add("DistID", SqlDbType.VarChar, 3).Value = sm.UserInfo.DistID
                    .Parameters.Add("CreateAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                End If
                If Me.ViewState("Action") = "EDIT" Then
                    .Parameters.Add("mtsn", SqlDbType.Decimal).Value = mtsn
                End If
                .ExecuteNonQuery()
            End With

        Catch ex As Exception
            OkErrorflag = False
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Finally
            '    objconn.Close()
            '    sqlCmd.Dispose()
        End Try

        Dim strScript As String = ""
        '若沒有錯誤狀況
        If OkErrorflag Then
            strScript = "<script language=""javascript"">" & vbCrLf

            If Me.ViewState("Action") = "ADD" Then
                strScript += "alert('會議日期及地點查詢建檔-新增成功!!');" & vbCrLf
            End If

            If Me.ViewState("Action") = "EDIT" Then
                strScript += "alert('會議日期及地點查詢建檔-修改成功!!');" & vbCrLf
            End If

            strScript += "location.href='OB_01_003.aspx?ID=" & Request("ID") & "';" & vbCrLf
            strScript += "</script>"

            Page.RegisterStartupScript("", strScript)
        End If
    End Sub

    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        TIMS.Utl_Redirect1(Me, "OB_01_003.aspx?ID=" & Request("ID"))
    End Sub

    Private Sub ddl_years_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_years.SelectedIndexChanged
        TIMS.Get_TenderCName(ddlTenderCName, sender.SelectedValue, objconn)
    End Sub
End Class
