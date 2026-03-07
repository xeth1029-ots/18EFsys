Partial Class SD_03_008
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'Dim conn As SqlConnection
        'TIMS.TestDbConn(Me, conn)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Syear = TIMS.GetSyear(Syear)
            Common.SetListItem(Syear, sm.UserInfo.Years)
            PlanID = TIMS.Get_LoginPlan(PlanID, Syear.SelectedValue, sm.UserInfo.DistID, objconn)

            DistID.Value = sm.UserInfo.DistID
            PlanID.SelectedValue = sm.UserInfo.PlanID

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

            Else '如果是委訓單位
                Button8.Visible = False
                center.Enabled = False
                Button4_Click(sender, e)
            End If

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
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim strErrmsg As String = ""
        Dim str As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        Dim ExportStr As String = "" '建立輸出文字
        Dim v_Syear As String = TIMS.GetListValue(Syear)
        If v_Syear = "" Then
            Common.MessageBox(Me, "請選擇年度!!")
            Exit Sub
        End If
        Dim v_PlanID As String = TIMS.GetListValue(PlanID)
        If v_PlanID = "" Then
            Common.MessageBox(Me, "請選擇訓練計畫!!")
            Exit Sub
        End If
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "請選擇訓練機構!!")
            Exit Sub
        End If
        Dim v_TYPE1 As String = TIMS.GetListValue(TYPE1)
        If v_TYPE1 = "" Then
            Common.MessageBox(Me, "請選擇匯出對象!!")
            Exit Sub
        End If

        If v_Syear <> "" Then
            str += " and pp.PlanYear = '" & v_Syear & "'"
        End If
        If v_PlanID <> "" Then
            str += " and pp.planID = '" & v_PlanID & "'"
        End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value <> "" Then
            str += " and cc.RID like '" & RIDValue.Value & "%'"
        End If
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value <> "" Then
            str += " and cc.OCID = '" & OCIDValue1.Value & "'"
        End If
        If v_TYPE1 = "2" Then   '在訓
            str += " and cc.FTDATE >= getdate() "
        ElseIf v_TYPE1 = "3" Then '結訓
            str += " and cc.FTDATE < getdate() "
        End If
        txtCJOB_NAME.Text = TIMS.ClearSQM(txtCJOB_NAME.Text)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If Me.txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then
            str += " and pp.CJOB_UNKEY='" & Me.cjobValue.Value & "'" & vbCrLf
        End If

        Dim sql As String = ""
        sql += " select ssd.name," & vbCrLf
        sql += " ssd.Email" & vbCrLf
        sql += " from class_classinfo cc" & vbCrLf
        sql += " join Class_StudentsOfClass cs on cc.ocid = cs.ocid" & vbCrLf
        sql += " join Stud_SubData ssd on ssd.sid = cs.sid" & vbCrLf
        sql += " join Plan_PlanInfo pp on pp.planid = cc.planid and pp.rid = cc.rid and pp.seqno = cc.seqno"
        sql += " where cc.notopen = 'N' AND  cs.StudStatus NOT IN (2,3) " & str
        'sql += " where cc.notopen = 'N' AND  cs.StudStatus NOT IN (2,3) and pp.Tplanid = '" & sm.UserInfo.TPlanID & "' and pp.planYear = '" & sm.UserInfo.Years & "'" & str
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無學員e-mail資料,不能匯出!!")
            Return
        End If

        '1:Outlook格式 2:Outlook Express格式
        Dim v_TYPE2 As String = TIMS.GetListValue(TYPE2)
        Select Case v_TYPE2
            Case "1", "2"
            Case Else
                Common.MessageBox(Me, "匯出格式有誤，請重新選擇!!")
                Return
        End Select

        '1:Outlook格式 2:Outlook Express格式
        Select Case v_TYPE2
            Case "1"
                'Outlook格式
                Dim MyFileName As String = String.Concat("SEmail", Now().ToString("yyyyMMddHHmmss"), ".xls")
                Dim MyPathFile As String = String.Concat("~\SD\03\Temp\", MyFileName)
                Dim MapPathF1 As String = Server.MapPath(MyPathFile)
                Const cst_SampleXLS As String = "~\SD\03\Sample3.xls"
                'copy一份sample資料---Start
                If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
                    Common.MessageBox(Me, "Sample檔案不存在")
                    Exit Sub
                End If
                Try
                    IO.File.Copy(Server.MapPath(cst_SampleXLS), MapPathF1, True)
                Catch ex As Exception
                    strErrmsg = ""
                    strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
                    strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
                    strErrmsg += ex.ToString & vbCrLf
                    Common.MessageBox(Me, strErrmsg)
                    'Exit Sub
                End Try

                '除去sample檔的唯讀屬性
                IO.File.SetAttributes(MapPathF1, IO.FileAttributes.Normal)
                'copy一份sample資料---------------------   End

                Using MyConn As New OleDb.OleDbConnection
                    MyConn.ConnectionString = TIMS.Get_OleDbStr(MapPathF1)
                    Try
                        MyConn.Open()
                    Catch
                        Common.MessageBox(Me, "Excel資料無法開啟連線!")
                        Exit Sub
                    End Try

                    For Each dr In dt.Rows
                        Dim Name2 As String = TIMS.ClearSQM(dr("name"))
                        Dim Email2 As String = TIMS.ClearSQM(dr("Email"))
                        sql = "INSERT INTO [Contacts] (名字,電子郵件地址)"
                        sql += "VALUES ('" & Name2.Replace("'", "") & "','" & Email2.Replace("'", "") & "')"

                        Using OleCmd As New OleDb.OleDbCommand(sql, MyConn)
                            Try
                                If MyConn.State = ConnectionState.Closed Then MyConn.Open()
                                OleCmd.ExecuteNonQuery()
                                'If conn.State = ConnectionState.Open Then conn.Close()
                            Catch ex As Exception
                                If MyConn.State = ConnectionState.Open Then MyConn.Close()
                                Throw ex
                            End Try
                        End Using
                    Next
                    If MyConn.State = ConnectionState.Open Then MyConn.Close()
                    '根據路徑建立資料庫連線，並取出學員資料填入---------------   End
                End Using

                '將新建立的excel存入記憶體下載-----   Start
                'Dim strErrmsg As String = ""
                strErrmsg = ""
                Try
                    Dim fr As New System.IO.FileStream(MapPathF1, IO.FileMode.Open)
                    Dim br As New System.IO.BinaryReader(fr)
                    Dim buf(fr.Length) As Byte

                    fr.Read(buf, 0, fr.Length)
                    fr.Close()

                    Response.Clear()
                    Response.ClearHeaders()
                    Response.Buffer = True
                    Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8))
                    Response.ContentType = "Application/vnd.ms-Excel"
                    'Common.RespWrite(Me, br.ReadBytes(fr.Length))
                    Response.BinaryWrite(buf)
                Catch ex As Exception
                    strErrmsg = ""
                    strErrmsg += "無法存取該檔案!!!" & vbCrLf
                    strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
                    strErrmsg += ex.ToString & vbCrLf
                Finally
                    '刪除Temp中的資料
                    If IO.File.Exists(MapPathF1) Then IO.File.Delete(MapPathF1)
                    If strErrmsg = "" Then Response.End()
                End Try
                If strErrmsg <> "" Then
                    Common.MessageBox(Me, strErrmsg)
                End If
                '將新建立的excel存入記憶體下載-----   End
            Case "2"
                'Outlook Express格式
                Dim s_OutlookCSV As String = String.Concat("SEmail", Now().ToString("yyyyMMddHHmmss"), ".csv")

                Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(s_OutlookCSV, System.Text.Encoding.UTF8))
                Response.ContentType = "Application/octet-stream"
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
                'ExportStr = "名字" & "," & "電子郵件地址" & ","
                ExportStr = "名字" & "," & "電子郵件地址"
                ExportStr += vbCrLf
                Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
                '建立資料面
                For Each dr In dt.Rows
                    ExportStr = ""
                    ExportStr = ExportStr & dr("name") & ","  '名字
                    'ExportStr = ExportStr & dr("Email") & "," '電子郵件地址
                    ExportStr = ExportStr & dr("Email")  '電子郵件地址
                    ExportStr += vbCrLf
                    Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
                Next
                Response.End()

            Case Else
                Common.MessageBox(Me, "匯出格式有誤，請重新選擇!!")
                Return
        End Select


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    Private Sub Syear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Syear.SelectedIndexChanged
        PlanID = TIMS.Get_LoginPlan(PlanID, Syear.SelectedValue, sm.UserInfo.DistID, objconn)
        center.Text = ""
        RIDValue.Value = ""
        TMID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        OCID1.Text = ""
    End Sub
End Class
