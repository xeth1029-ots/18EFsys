Partial Class SD_05_005_R
    Inherits AuthBasePage

    Dim s_log1 As String = ""
    Dim CheckMode As String = ""
    Dim PageSize As Integer = 24   '分頁-- 每頁資料 筆數
    Dim CourseNumSize As Integer = 4 '每頁科目顯示數目

    'Dim int_BasicPoint As String = ""
    'Dim chk1, chk2, chk3, chk4, chk5, chk6, chk7, chk8, chk9, chk10, chk11, chk12 As String
    Dim int_BasicPoint As Integer = 0

    Dim in_SOCID As String = ""
    Dim OCID As String = ""
    Dim RIDvalue As String = ""
    Dim ChooseClass As String = ""
    Dim rqSignType As String = ""

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁    
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        OCID = TIMS.ClearSQM(Request("OCID"))
        RIDvalue = TIMS.ClearSQM(Request("RID"))
        in_SOCID = TIMS.CombiSQM2IN(Request("SOCID"))
        ChooseClass = TIMS.ClearSQM(Request("ChooseClass"))
        int_BasicPoint = Val(TIMS.ClearSQM(Request("BasicPoint")))
        rqSignType = TIMS.ClearSQM(Request("SignType"))

        '設定頁數
        CourseNumSize = SetCourseNumSize()

        If Not IsPostBack Then
            '取得分頁頁數、單位、計劃名稱
            'GetRecordSize()
            'CreateRpt(1, 1) '只列出第一頁
            '列出所有分頁資料
            Call PrintResult()
            'If tb_PageTotal.Text = "1" Then
            '    printNote()
            'End If
        End If
    End Sub

    Private Sub PrintResult(Optional ByVal iPageIndex As Integer = 0)
        '取得分頁頁數、單位、計劃名稱
        Call GetRecordSize()

        'Dim x As Integer = 0
        'Dim y As Integer = 0
        Dim iPageEnd As Integer = 0
        Dim iCoursePageNum As Integer = 0
        Dim iAllPage As Integer = 0

        '取出目前選取科目 /CourseNumSize('每頁科目顯示數目)
        iAllPage = Math.Ceiling(GetCourseCount() / CourseNumSize)

        If iAllPage = 0 Then iAllPage = 1
        If tb_PageTotal.Text = "0" Then tb_PageTotal.Text = "1"
        If CInt(Val(tb_PageTotal.Text)) <= 0 Then tb_PageTotal.Text = "1"

        '每頁可呈現之科目欄位數
        iCoursePageNum = Math.Ceiling(CInt(Val(tb_Page.Text)) / (CInt(Val(tb_PageTotal.Text)) / iAllPage))

        If iPageIndex <> 0 Then
            iPageEnd = 1
        Else
            iPageEnd = Convert.ToInt16(Me.ViewState("PageCount"))
        End If
        For y As Integer = 1 To Convert.ToInt16(Me.ViewState("PageCount"))
            For x As Integer = 1 To iAllPage
                '列印資料狀況
                Call CreateRpt(y, x)
                If Convert.ToInt16(Me.ViewState("FinalPageSize")) < 15 Then
                    If Not (y = Convert.ToInt16(Me.ViewState("PageCount")) And x = iAllPage) Then
                        '列印每頁空白資料列
                        Call PrintEmptPage()    '分頁
                    End If
                Else
                    '列印每頁空白資料列
                    Call PrintEmptPage() '分頁
                End If
            Next
        Next
        '列印備註
        Call printNote()        '列印備註
        doPrint.Value = "Y"
        WaterPage.Value = tb_PageTotal.Text
    End Sub

    '列印每頁空白資料列
    Sub PrintEmptPage()
        Dim nl As HtmlGenericControl = New HtmlGenericControl
        print_content.Controls.Add(nl)
        nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
    End Sub

    '列印備註
    Sub printNote()
        Dim rptTb As HtmlTable
        Dim rptRow As HtmlTableRow
        Dim rptCell As HtmlTableCell
        Dim nl As HtmlGenericControl = New HtmlGenericControl
        print_content.Controls.Add(nl)
        nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;'>"
        '* --備註 --*
        rptTb = New HtmlTable
        rptTb.Attributes.Add("style", "BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
        rptTb.Attributes.Add("align", "left")
        rptTb.Attributes.Add("border", 0)
        print_content.Controls.Add(rptTb)

        'Dim m As Integer = 1
        'Dim n As Integer = 1

        For m As Integer = 1 To 5      '(一) 共6個row
            rptRow = New HtmlTableRow  '1 (一)
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
            rptCell.Attributes.Add("colspan", 12)
            Select Case m
                Case 1
                    rptCell.InnerHtml = "設定課程群組關係："
                Case 2
                    rptCell.InnerHtml = "&nbsp"
                Case 3
                    rptCell.InnerHtml = "主課程(30)和細課程(10,10,10)都可以排課,在記錄成績選(用主課程來選,主課程+細課程排課時數)□時數多少<br/>以上列出來,有選進來的課程,才須記錄成績時並且列印,成績只打在主課程裡面。"
                Case 4
                    rptCell.InnerHtml = "&nbsp"
                Case 5
                    rptCell.InnerHtml = Me.ViewState("recordNote")
            End Select
        Next
        For m As Integer = 1 To 3     '(二)共三個row
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            For n As Integer = 1 To 6   '每個row各有6個column
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                rptCell.Attributes.Add("colspan", 1)
                Select Case m
                    Case 1
                        'If m = 1 Then
                        Select Case n
                            Case 1
                                rptCell.InnerHtml = "&nbsp"
                            Case 2
                                rptCell.InnerHtml = "&nbsp"
                            Case 3
                                rptCell.InnerHtml = "&nbsp"
                            Case 4
                                rptCell.InnerHtml = "&nbsp"
                            Case 5
                                rptCell.InnerHtml = "&nbsp"
                            Case 6
                                rptCell.Attributes.Add("colspan", 7)
                                rptCell.InnerHtml = "&nbsp"
                        End Select
                        'ElseIf m = 2 Then
                    Case 2
                        Select Case rqSignType
                            Case "1"
                                Select Case n
                                    Case 1
                                        rptCell.InnerHtml = IIf(Request("chk1") = "1", "導師:", "")
                                        'rptCell.InnerHtml = "&nbsp"
                                    Case 2
                                        rptCell.InnerHtml = "&nbsp"
                                    Case 3
                                        rptCell.InnerHtml = IIf(Request("chk2") = "1", "教學行政股長:", "")
                                        'rptCell.InnerHtml = "&nbsp"
                                    Case 4
                                        rptCell.InnerHtml = "&nbsp"
                                    Case 5
                                        rptCell.InnerHtml = IIf(Request("chk3") = "1", "批示:", "")
                                        'rptCell.InnerHtml = "&nbsp"
                                    Case 6
                                        rptCell.Attributes.Add("colspan", 7)
                                        rptCell.InnerHtml = "&nbsp"
                                End Select
                            Case Else '"2"
                                Select Case n
                                    Case 1
                                        rptCell.InnerHtml = IIf(Request("chk10") = "1", "承辦: ", "")
                                        'rptCell.InnerHtml = "&nbsp"
                                    Case 2
                                        rptCell.InnerHtml = "&nbsp"
                                    Case 3
                                        rptCell.InnerHtml = IIf(Request("chk11") = "1", "主管:", "")
                                        'rptCell.InnerHtml = "&nbsp"
                                    Case 4
                                        rptCell.InnerHtml = "&nbsp"
                                    Case 5
                                        rptCell.InnerHtml = IIf(Request("chk12") = "1", "批示:", "")
                                        'rptCell.InnerHtml = "&nbsp"
                                    Case 6
                                        rptCell.Attributes.Add("colspan", 7)
                                        rptCell.InnerHtml = "&nbsp"
                                End Select
                        End Select

                    Case 3
                        'ElseIf m = 3 Then
                        Select Case n
                            Case 1
                                rptCell.InnerHtml = IIf(Request("chk4") = "1", "股長:", "")
                                'rptCell.InnerHtml = "&nbsp"
                            Case 2
                                rptCell.InnerHtml = "&nbsp"
                            Case 3
                                rptCell.InnerHtml = IIf(Request("chk5") = "1", "科長:", "")
                                'rptCell.InnerHtml = "&nbsp"
                            Case 4
                                rptCell.InnerHtml = "&nbsp"
                            Case 5
                                rptCell.InnerHtml = "&nbsp"
                            Case 6
                                rptCell.Attributes.Add("colspan", 7)
                                rptCell.InnerHtml = "&nbsp"
                        End Select
                        'End If
                End Select
            Next
            rptRow = New HtmlTableRow  '強制跳行
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
            rptCell.Attributes.Add("colspan", 12)
            rptCell.InnerHtml = "&nbsp"
        Next
    End Sub

    Function SET_PARMS1(ByRef iGVID As Integer, ByRef sm As SessionModel, ByRef ItemVar As Integer, ByRef oResultTyp As Object) As Hashtable
        'Dim GVID As String = TIMS.GetMyValue2(HtP, "GVID")
        'Dim DistID As String = TIMS.GetMyValue2(HtP, "DistID")
        'Dim TPlanID As String = TIMS.GetMyValue2(HtP, "TPlanID")
        'Dim ItemVar As Integer = TIMS.GetMyValue2(HtP, "ItemVar") 'ItemVar 1:'學科百分比/2:術科百分比
        'Dim ResultTyp As String = TIMS.GetMyValue2(HtP, "ResultTyp")
        Dim s_DISTID As String = sm.UserInfo.DistID
        'Dim s_TPlanID As String = sm.UserInfo.TPlanID
        If sm.UserInfo.LID = 0 AndAlso RIDvalue <> "" Then
            '署-重新設計DISTID
            s_DISTID = TIMS.Get_DistID_RID(RIDvalue, objconn)
        End If

        Dim parms As New Hashtable From {
            {"GVID", iGVID},
            {"DistID", s_DISTID},
            {"TPlanID", sm.UserInfo.TPlanID},
            {"ItemVar", ItemVar},
            {"ResultTyp", If(oResultTyp Is Nothing, "", Convert.ToString(oResultTyp))}
        }
        Return parms
    End Function

    '設定頁數
    Function SetCourseNumSize() As Integer
        Dim iRst As Integer = 5

        Dim s_ResultTyp As String = If(ViewState("ResultTyp") IsNot Nothing, Convert.ToString(ViewState("ResultTyp")), "")
        Dim ResultTyp As Double = GetPara(objconn, SET_PARMS1(13, sm, 1, s_ResultTyp)) '成績計算方式 ( 1, 各科平均法  2, 訓練時數權重法 )
        Dim percent1 As Double = GetPara(objconn, SET_PARMS1(17, sm, 1, s_ResultTyp)) '學科百分比
        Dim percent2 As Double = GetPara(objconn, SET_PARMS1(17, sm, 2, s_ResultTyp)) '術科百分比
        If percent1 <> 0 AndAlso percent2 <> 0 AndAlso ResultTyp = 1 Then iRst = 3

        Return iRst
    End Function

    '取得分頁頁數、單位、計劃名稱
    Private Sub GetRecordSize()
        Dim sql As String = ""
        Dim sql_2 As String = ""
        Dim sql_3 As String = ""
        Dim sql_4 As String = ""
        Dim dr As DataRow
        'Dim da As New SqlDataAdapter
        Dim dt As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim dt4 As New DataTable
        'Dim conn As New SqlConnection
        Dim OCIDValue1 As String = TIMS.ClearSQM(Request("OCID"))
        Dim SOCID As String = TIMS.ClearSQM(Request("SOCID"))
        If OCIDValue1 = "" OrElse SOCID = "" Then Exit Sub

        SOCID = TIMS.ClearSQM(TIMS.CombiSQM2IN(SOCID))
        If SOCID = "" Then Exit Sub

        If ChooseClass = "" Then Exit Sub
        ChooseClass = TIMS.ClearSQM(TIMS.CombiSQM2IN(ChooseClass))
        If ChooseClass = "" Then Exit Sub

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1, objconn)
        If drCC Is Nothing Then Exit Sub

        '取出本班所有學員 
        sql = ""
        sql &= " SELECT concat(a.NAME,case when b.STUDSTATUS in (2,3) then '(*)' end) Name" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(b.StudentID) StudentID" & vbCrLf
        sql &= " ,b.SOCID,b.OCID,b.StudStatus,b.CloseDate,c.FTDate" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO c" & vbCrLf
        sql &= " JOIN ID_CLASS d ON d.clsid=c.clsid" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON b.OCID=c.OCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO a on a.SID =b.SID" & vbCrLf
        sql &= " WHERE c.OCID='" & OCIDValue1 & "'" & vbCrLf
        sql &= " Order By b.StudentID" & vbCrLf

        If CInt(sm.UserInfo.Years) < Now.Year Then
            '單位名稱變更歷史資料
            sql_2 = ""
            sql_2 &= " select o2.years ,o2.orgname,o2.editid" & vbCrLf
            sql_2 &= " FROM CLASS_CLASSINFO c1" & vbCrLf
            sql_2 &= " JOIN ORG_ORGINFO o1 on c1.comidno = o1.comidno" & vbCrLf
            sql_2 &= " JOIN ID_PLAN d on c1.PlanID =d.PlanID" & vbCrLf
            sql_2 &= " JOIN KEY_PLAN e on d.TPlanID=e.TPlanID" & vbCrLf
            sql_2 &= " LEFT JOIN ORG_ORGNAMEHISTORY o2 on o1.orgid = o2.orgid" & vbCrLf
            sql_2 &= " where c1.ocid='" & OCIDValue1 & "'" & vbCrLf
            sql_2 &= " and o2.years=d.years" & vbCrLf
            sql_2 &= " order by o2.editid desc" & vbCrLf
        Else
            '登入年度為本年度則直接抓 org_orginfo 內的orgname
            sql_2 = ""
            sql_2 &= " select CONVERT(varchar, o1.ModifyDate, 111) years"
            sql_2 &= " ,c1.ocid ,o1.orgid ,o1.orgname"
            sql_2 &= " from class_classinfo c1"
            sql_2 &= " join org_orginfo o1 on c1.comidno = o1.comidno"
            sql_2 &= " where c1.ocid='" & OCIDValue1 & "'"
        End If

        sql_3 = ""
        sql_3 &= " select CONVERT(varchar, o1.ModifyDate, 111) ModifyYear"
        sql_3 &= " ,o1.orgname ,e.PlanName"
        sql_3 &= " from class_classinfo c1"
        sql_3 &= " join org_orginfo o1 on c1.comidno = o1.comidno"
        sql_3 &= " left join ID_Plan d on c1.PlanID =d.PlanID"
        sql_3 &= " left join key_plan e on d.TPlanID=e.TPlanID"
        sql_3 &= " where c1.ocid ='" & OCIDValue1 & "'"

        '取出所有學員成績    
        sql_4 = ""
        sql_4 &= " SELECT SOCID,COURID" & vbCrLf
        sql_4 &= " FROM Stud_TrainingResults" & vbCrLf
        sql_4 &= " WHERE CourID in (" & ChooseClass & ")" & vbCrLf
        sql_4 &= " AND SOCID in (" & SOCID & ")  order by SOCID"

        'conn = DbAccess.GetConnection()
        'conn.Open()
        Call TIMS.OpenDbConn(objconn)
        Dim sCmd1 As New SqlCommand(sql, objconn)
        With sCmd1
            dt.Load(.ExecuteReader())
        End With

        Dim sCmd2 As New SqlCommand(sql_2, objconn)
        With sCmd2
            dt2.Load(.ExecuteReader())
        End With

        Dim sCmd3 As New SqlCommand(sql_3, objconn)
        With sCmd3
            dt3.Load(.ExecuteReader())
        End With

        Dim sCmd4 As New SqlCommand(sql_4, objconn)
        With sCmd4
            dt4.Load(.ExecuteReader())
        End With

        If dt.Rows.Count > 0 Then
            Me.ViewState("PerPageSize") = dt4.Rows.Count.ToString                  '每頁資料筆數(總資料)
            Me.ViewState("PerPageEmptSize") = (PageSize - dt.Rows.Count).ToString '每頁空白資料筆數
            Me.ViewState("PageCount") = Convert.ToString(Math.Ceiling(Convert.ToInt16(dt.Rows.Count) / PageSize)) '總頁數 /每頁頁數
            Me.ViewState("FinalPageSize") = Convert.ToString(Convert.ToInt16(dt.Rows.Count) Mod PageSize)         '最後一頁的頁數
            tb_PageTotal.Text = Convert.ToString(Convert.ToInt16(Me.ViewState("PageCount")) * Math.Ceiling(GetCourseCount() / CourseNumSize))      '總頁數*科目循環=總分頁數
            tb_Page.Text = "1"
        End If
        If dt2.Rows.Count > 0 Then
            If dt3.Rows.Count > 0 Then
                If Convert.ToString(dt3.Rows(0)("ModifyYear")) = Convert.ToString(dt2.Rows(0)("years")) Then  '如 org_orginfo  最後變更年度為 登入計畫年度則取目前 org_orginfo 之 orgname
                    Me.ViewState("OrgName") = Convert.ToString(dt3.Rows(0)("orgname")) '機構名稱
                Else
                    Me.ViewState("OrgName") = Convert.ToString(dt2.Rows(0)("orgname")) '機構名稱
                End If
            End If
        Else
            If dt3.Rows.Count > 0 Then
                dr = dt3.Rows(0)
                Me.ViewState("OrgName") = Convert.ToString(dr("orgname")) '機構名稱
                '登入年度
                Me.ViewState("PlanName") = sm.UserInfo.Years & "  年度 " & Convert.ToString(dr("PlanName")) & "  學員成績表"
            End If
        End If

    End Sub

#Region "NO USE"
    'Function Get_ChooseClass(ByVal OCIDValue1 As String) As String
    '    Get_ChooseClass = ""
    '    Dim sql As String
    '    Dim dt As DataTable
    '    sql = "" & vbCrLf
    '    sql &= " SELECT Distinct b.CourseName, a.CourID FROM" & vbCrLf
    '    sql &= " (SELECT * FROM Stud_TrainingResults WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1 & "')) a" & vbCrLf
    '    sql &= " LEFT JOIN Course_CourseInfo b ON a.CourID=b.CourID order by b.courseName" & vbCrLf
    '    dt = DbAccess.GetDataTable(sql, objconn)
    '    If dt.Rows.Count > 0 Then
    '        For Each dr As DataRow In dt.Rows
    '            If Get_ChooseClass = "" Then
    '                Get_ChooseClass = "'" & dr("CourID") & "'"
    '            Else
    '                Get_ChooseClass += ",'" & dr("CourID") & "'"
    '            End If
    '        Next
    '    End If
    'End Function

#End Region


    Public Shared Function GetPara(ByRef oConn As SqlConnection, ByRef HtP As Hashtable) As Double  '取得參數設定值
        'Optional ByVal ItemVar As Integer = 1
        'ByVal GVID As String, ByVal DistID As String, ByVal TPlanID As String, ByVal ItemVar As Integer
        Dim GVID As String = TIMS.GetMyValue2(HtP, "GVID")
        Dim DistID As String = TIMS.GetMyValue2(HtP, "DistID")
        Dim TPlanID As String = TIMS.GetMyValue2(HtP, "TPlanID")
        Dim ItemVar As Integer = TIMS.GetMyValue2(HtP, "ItemVar") 'ItemVar 1:'學科百分比/2:術科百分比
        Dim ResultTyp As String = TIMS.GetMyValue2(HtP, "ResultTyp")

        Dim i_rst As Double = 0
        Dim sql As String = ""
        sql = "SELECT * FROM SYS_GLOBALVAR WITH(NOLOCK) WHERE DistID=@DistID and TPlanID=@TPlanID"
        Dim sCmd As New SqlCommand(sql, oConn)
        TIMS.OpenDbConn(oConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("DistID", SqlDbType.VarChar).Value = DistID
            .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID
            dt.Load(.ExecuteReader())
        End With

        'dt = DbAccess.GetDataTable(sql, oConn)
        Select Case GVID
            Case "3" '操行底分
                If dt.Select("GVID='3'").Length <> 0 Then
                    i_rst = dt.Select("GVID='3'")(0)("ItemVar1")
                End If
            Case "13"  '成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
                'Me.ViewState("ResultTyp")
                i_rst = 1 '預設取各科平均法
                If (ResultTyp = "") AndAlso dt.Select("GVID='13'").Length <> 0 Then
                    '預設取各科平均法
                    Dim s_ItemVar1_TMP As String = dt.Select("GVID='13'")(0)("ItemVar1")
                    If TIMS.IsNumeric1(s_ItemVar1_TMP) Then i_rst = Val(s_ItemVar1_TMP)    '預設取各科平均法
                End If
            Case "17"      '學術科百分比 (學科ItemVar1 ex:0.4,術科ItemVar2  ex:0.4) 
                If dt.Select("GVID='17'").Length = 0 Then         '預設各為50%
                    i_rst = 0 'ItemVar 1:'學科百分比/2:術科百分比
                Else
                    i_rst = 0 'ItemVar 1:'學科百分比/2:術科百分比
                    Select Case ItemVar
                        Case 1
                            i_rst = Math.Round(Convert.ToDouble(dt.Select("GVID='17'")(0)("ItemVar1")), 2)
                        Case 2
                            i_rst = Math.Round(Convert.ToDouble(dt.Select("GVID='17'")(0)("ItemVar2")), 2)
                    End Select
                End If
        End Select
        Return i_rst
    End Function

    '取出目前選取科目 
    Private Function GetCourseCount() As Integer
        Dim rst As Integer = 0

        '先取出目前所選取的科目  
        Dim sql_1 As String = ""
        sql_1 &= " SELECT CourID,CourseName "
        sql_1 &= " ,case  Classification1 when 1 then '學科' when 2 then '術科' end classType "
        sql_1 &= " ,Classification1 "
        sql_1 &= " ,0 TotalHours"
        sql_1 &= " FROM Course_CourseInfo"
        sql_1 &= " WHERE CourID in (" & ChooseClass & " ) "
        sql_1 &= " Order By CourseName "
        Call TIMS.OpenDbConn(objconn)
        Dim sCmd As New SqlCommand(sql_1, objconn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count > 0 Then
            rst = dt1.Rows.Count
        End If
        Return rst
    End Function

    Private Function getStudResult(ByVal Item As Integer, Optional ByVal OneSocid As String = "") As DataTable '090825
        Dim sql_1 As String = ""
        Dim sql_1_2 As String = ""
        Dim sql_2 As String = ""
        Dim sql_3 As String = ""
        Dim sql_A As String = ""
        Dim sql_B As String = ""
        Dim sql_C As String = ""
        'Dim ds As New DataSet

        Dim dt1 As New DataTable
        Dim dt1_2 As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable

        Dim dtReult1 As New DataTable
        Dim dtReult2 As New DataTable
        Dim dtReult3 As New DataTable 'nothing

        'Dim dr As DataRow
        'Dim dr1 As DataRow
        'Dim dr2 As DataRow
        'Dim dr3 As DataRow
        'Dim dr4 As DataRow

        OCID = TIMS.ClearSQM(OCID)

        '先取出目前所選取的科目  
        sql_1 = " SELECT COURID,COURSENAME "
        sql_1 &= " ,CASE CLASSIFICATION1 WHEN 1 THEN '學科' WHEN 2 THEN '術科' END CLASSTYPE ,CLASSIFICATION1 ,CONVERT([numeric] (10,0),NULL) TOTALHOURS" & vbCrLf
        sql_1 &= " FROM COURSE_COURSEINFO" & vbCrLf
        sql_1 &= String.Concat(" WHERE COURID IN (", ChooseClass, ")") & vbCrLf
        sql_1 &= " ORDER BY COURSENAME "

        '檢查學員成績的有登錄成績的科目總數    
        sql_1_2 = ""
        sql_1_2 &= " SELECT COURID FROM STUD_TRAININGRESULTS" & vbCrLf
        sql_1_2 &= String.Concat(" WHERE COURID IN (", ChooseClass, ")") & vbCrLf
        sql_1_2 &= String.Concat(" AND SOCID in (", in_SOCID, ")")
        sql_1_2 &= " GROUP BY COURID ORDER BY COURID" & vbCrLf

        '取出所有學員成績    
        sql_2 = "" & vbCrLf
        sql_2 &= " SELECT CONVERT([numeric] (10,0),NULL) OCID,SOCID,COURID" & vbCrLf
        sql_2 &= " ,CONVERT(NVARCHAR(200),NULL) COURSENAME,RESULTS,CONVERT(VARCHAR(33),NULL) CLASSTYPE" & vbCrLf
        sql_2 &= " ,CONVERT(VARCHAR(1),NULL) CLASSIFICATION1,CONVERT([numeric] (10,0),NULL) HOURS,CONVERT([numeric] (10,0),NULL) RESULTS2" & vbCrLf
        sql_2 &= " FROM STUD_TRAININGRESULTS" & vbCrLf
        sql_2 &= String.Concat(" WHERE COURID IN (", ChooseClass, ")") & vbCrLf
        sql_2 &= If(OneSocid <> "", String.Concat(" AND SOCID IN (", OneSocid, ")"), If(in_SOCID <> "", String.Concat(" AND SOCID IN (", in_SOCID, ")"), ""))
        sql_2 &= " ORDER BY SOCID,COURSENAME" & vbCrLf

        '產生一個空的table
        sql_A = "" & vbCrLf
        sql_A &= " SELECT COURID,COURSENAME"
        sql_A &= " ,CONVERT(VARCHAR(33),NULL) CLASSTYPE ,CLASSIFICATION1,CONVERT([numeric] (10,0),NULL) TOTALHOURS" & vbCrLf
        sql_A &= " FROM COURSE_COURSEINFO WHERE 1<>1"

        '產生一個空的table
        sql_B = "" & vbCrLf
        sql_B &= " SELECT CONVERT([numeric] (10,0),NULL) OCID,SOCID,COURID" & vbCrLf
        sql_B &= " ,CONVERT(NVARCHAR(200),NULL) COURSENAME,RESULTS,CONVERT(VARCHAR(33),NULL) CLASSTYPE" & vbCrLf
        sql_B &= " ,CONVERT(VARCHAR(1),NULL) CLASSIFICATION1,CONVERT([numeric] (10,0),NULL) HOURS,CONVERT([numeric] (10,0),NULL) RESULTS2" & vbCrLf
        sql_B &= " FROM STUD_TRAININGRESULTS WHERE 1<>1" & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim sCmd_1 As New SqlCommand(sql_1, objconn)
        With sCmd_1
            dt1.Load(.ExecuteReader())
        End With

        Dim sCmd_1_2 As New SqlCommand(sql_1_2, objconn)
        With sCmd_1_2
            dt1_2.Load(.ExecuteReader())
        End With

        Dim sCmd_2 As New SqlCommand(sql_2, objconn)
        With sCmd_2
            dt2.Load(.ExecuteReader())
            TIMS.CHG_dtReadOnly(dt2)
        End With

        sql_3 = " SELECT * FROM CLASS_SCHEDULE WHERE OCID=@OCID" '取出排課檔  
        Dim sCmd_3 As New SqlCommand(sql_3, objconn)
        With sCmd_3
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(OCID)
            dt3.Load(.ExecuteReader())
        End With

        Dim sCmd_A As New SqlCommand(sql_A, objconn)
        With sCmd_A
            dtReult1.Load(.ExecuteReader())
            TIMS.CHG_dtReadOnly(dtReult1)
        End With

        Dim sCmd_B As New SqlCommand(sql_B, objconn)
        With sCmd_B
            dtReult2.Load(.ExecuteReader())
            TIMS.CHG_dtReadOnly(dtReult2)
        End With

        Dim CountHours As Integer = 0
        For Each dr1 As DataRow In dt1.Rows       '列出所有選取的科目
            Dim i_TotalHours As Integer = If(Convert.ToString(dr1("TotalHours")) <> "", Val(dr1("TotalHours")), 0)
            For Each dr3 As DataRow In dt3.Rows   '依序列出排課
                For i As Integer = 1 To 12
                    Dim colNMClass1 As String = String.Concat("Class", i)
                    If Not IsDBNull(dr3(colNMClass1)) AndAlso Convert.ToString(dr3(colNMClass1)) <> "" Then
                        If Convert.ToString(dr3(colNMClass1)) = Convert.ToString(dr1("CourID")) Then
                            i_TotalHours += 1
                        End If
                    End If
                Next
            Next
            dt1.Columns("TotalHours").ReadOnly = False
            dr1("TotalHours") = i_TotalHours

            Dim dr As DataRow = dtReult1.NewRow  '填入空的table
            dr("CourID") = dr1("CourID")
            dr("CourseName") = dr1("CourseName")
            dr("classType") = dr1("classType")
            dr("Classification1") = dr1("Classification1")
            dr("TotalHours") = dr1("TotalHours")
            dtReult1.Rows.Add(dr)
        Next

        For Each dr2 As DataRow In dt2.Rows       '列出所有選取的學員 
            For Each dr4 As DataRow In dtReult1.Rows
                If Convert.ToString(dr2("COURID")) = Convert.ToString(dr4("COURID")) Then
                    dr2("COURID") = dr4("COURID")
                    dr2("COURSENAME") = dr4("COURSENAME")
                    dr2("CLASSTYPE") = dr4("CLASSTYPE")
                    dr2("CLASSIFICATION1") = dr4("CLASSIFICATION1")
                    dr2("HOURS") = dr4("TotalHours")
                    dr2("RESULTS2") = dr2("RESULTS") * dr4("TotalHours")

                    Dim dr As DataRow = dtReult2.NewRow
                    dr("OCID") = dr2("OCID")
                    dr("SOCID") = dr2("SOCID")
                    dr("COURID") = dr2("COURID")
                    dr("COURSENAME") = dr2("COURSENAME")
                    dr("RESULTS") = dr2("RESULTS")
                    dr("CLASSTYPE") = dr2("CLASSTYPE")
                    dr("CLASSIFICATION1") = dr2("CLASSIFICATION1")
                    dr("HOURS") = dr2("HOURS")
                    dr("RESULTS2") = dr2("RESULTS2")
                    dtReult2.Rows.Add(dr)
                End If
            Next
        Next

        Select Case Item
            Case 1
                If dt1.Rows.Count >= 1 AndAlso dt1_2.Rows.Count = 1 Then '所選取的科目大於一個且 學員成績檔中只有登錄一個科目成績時 -->以「各科平均法」計算
                    Me.ViewState("ResultTyp") = "1"
                End If
                Return dt1
            Case 2
                Return dtReult2
        End Select

        Return dtReult3
    End Function

    '列印資料狀況 (Main)
    Private Sub CreateRpt(ByVal PageNum As Integer, Optional ByVal CoursePageNum As Integer = 1)
        Dim rptTb As HtmlTable
        Dim rptRow As HtmlTableRow
        Dim rptCell As HtmlTableCell

        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim dt4 As New DataTable
        Dim dt5 As New DataTable

        Dim dr1 As DataRow
        'Dim da As New SqlDataAdapter
        'Dim conn As New SqlConnection
        Dim sql_1 As String = ""
        Dim sql_2 As String = ""
        'Dim sql_3 As String = ""
        'Dim sql_4 As String = ""
        Dim sql_5 As String = ""
        Dim sqlstr As String = ""

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim recordNote As String = ""

        Dim s_ResultTyp As String = If(ViewState("ResultTyp") IsNot Nothing, Convert.ToString(ViewState("ResultTyp")), "")
        Dim ResultTyp As Double = GetPara(objconn, SET_PARMS1(13, sm, 1, s_ResultTyp)) '成績計算方式 ( 1, 各科平均法  2, 訓練時數權重法 )
        Dim percent1 As Double = GetPara(objconn, SET_PARMS1(17, sm, 1, s_ResultTyp)) '學科百分比
        Dim percent2 As Double = GetPara(objconn, SET_PARMS1(17, sm, 2, s_ResultTyp)) '術科百分比
        'Dim ResultTyp As Integer = GetPara(13, Me.sm.UserInfo.DistID, Me.sm.UserInfo.TPlanID)       '成績計算方式 ( 1, 各科平均法  2, 訓練時數權重法 )
        'Dim percent1 As Double = GetPara(17, Me.sm.UserInfo.DistID, Me.sm.UserInfo.TPlanID, 1)      '學科百分比
        'Dim percent2 As Double = GetPara(17, Me.sm.UserInfo.DistID, Me.sm.UserInfo.TPlanID, 2)      '術科百分比

        Dim ChooseClass As String = TIMS.ClearSQM(Request("ChooseClass"))
        Dim OCIDValue1 As String = TIMS.ClearSQM(Request("OCID"))
        RIDvalue = TIMS.ClearSQM(Request("RID"))
        'Dim SOCID As String = TIMS.ClearSQM(Request("SOCID"))
        Dim SOCID_in As String = TIMS.CombiSQM2IN(Request("SOCID"))
        Dim PrintDataTyp As String = TIMS.ClearSQM(Request("PrintDataTyp"))  '報表列印格式 1:空白 2:有帶資料

        '取出目前選取科目 
        dt1 = getStudResult(1)
        '建立資料 Start
        dt2 = getStudResult(2)

        '取出學號、姓名
        Dim sql_3 As String = ""
        sql_3 &= " SELECT a.Name  + case when b.studstatus in (2,3) then '(*)' else '' end Name" & vbCrLf
        sql_3 &= " ,dbo.FN_CSTUDID2(b.STUDENTID) StudentID" & vbCrLf
        sql_3 &= " ,b.SOCID, b.OCID,b.StudStatus" & vbCrLf
        sql_3 &= " ,b.CloseDate, c.FTDate" & vbCrLf
        sql_3 &= " FROM STUD_STUDENTINFO a" & vbCrLf
        sql_3 &= " JOIN CLASS_STUDENTSOFCLASS b ON b.SID=a.SID" & vbCrLf
        sql_3 &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID" & vbCrLf
        sql_3 &= " JOIN ID_CLASS d ON c.clsid = d.clsid" & vbCrLf
        sql_3 &= " WHERE c.OCID=@OCID" & vbCrLf
        sql_3 &= If(SOCID_in <> "", String.Concat(" AND b.SOCID IN (", SOCID_in, ")"), " AND 1<>1") & vbCrLf
        sql_3 &= " ORDER BY b.StudentID" & vbCrLf

        'Dim int_BasicPoint As Integer = 0
        'int_BasicPoint = GetPara(3, sm.UserInfo.DistID, sm.UserInfo.TPlanID)
        If int_BasicPoint = 0 Then
            'Dim s_ResultTyp As String = If(ViewState("ResultTyp") IsNot Nothing, Convert.ToString(ViewState("ResultTyp")), "")
            int_BasicPoint = GetPara(objconn, SET_PARMS1(3, sm, 1, s_ResultTyp))
            'int_BasicPoint = GetPara(3, sm.UserInfo.DistID, sm.UserInfo.TPlanID)
        End If
        s_log1 = String.Format("##SD_05_005_R CreateRpt ,int_BasicPoint: {0}", int_BasicPoint)
        TIMS.LOG.Debug(s_log1)

        '名次
        OCIDValue1 = TIMS.ClearSQM(OCIDValue1)

        Dim dtAvg As New DataTable
        Dim drAvg As DataRow = Nothing

        Call TIMS.OpenDbConn(objconn)
        s_log1 = String.Format("##SD_05_005_R SD_05_005.Get_Stud_Turnout(sm.UserInfo.PlanID:{0}, RIDvalue:{1}, OCIDValue1:{2},, int_BasicPoint:{3})", sm.UserInfo.PlanID, RIDvalue, OCIDValue1, int_BasicPoint)
        TIMS.LOG.Debug(s_log1)
        dt4 = SD_05_005.Get_Stud_Turnout(sm.UserInfo.PlanID, RIDvalue, OCIDValue1, objconn, int_BasicPoint)

        Dim sCmd_3 As New SqlCommand(sql_3, objconn)
        With sCmd_3
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(OCIDValue1)
            dt3.Load(.ExecuteReader())
        End With

        sql_5 = " SELECT SOCID,OCID,ISNULL(RANK,0) RANK,STUDSTATUS FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID"
        Dim sCmd_5 As New SqlCommand(sql_5, objconn)
        With sCmd_5
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(OCIDValue1)
            dt5.Load(.ExecuteReader())
        End With

        '將所有學員計算之總平均取出
        'sqlstr = "select  ''  as  socid ,0.0 as  totalAvg  "
        sqlstr = " select convert(varchar(100),'') as socid,0.0 as totalAvg"
        Dim sCmd_Q As New SqlCommand(sqlstr, objconn)
        With sCmd_Q
            dtAvg.Load(.ExecuteReader())
            TIMS.CHG_dtReadOnly(dtAvg)
            TIMS.CHG_dtAllowDBNull(dtAvg)
        End With

        rptTb = Nothing
        ' -   表首  start   
        rptTb = New HtmlTable
        rptTb.Attributes.Add("style", "BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
        rptTb.Attributes.Add("align", "center")
        rptTb.Attributes.Add("border", 0)
        print_content.Controls.Add(rptTb)
        ' ----------
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)

        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:16pt;font-family:DFKai-SB")
        rptCell.Attributes.Add("colspan", 12)
        rptCell.InnerHtml = Me.ViewState("OrgName")      '單位

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)

        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:16pt;font-family:DFKai-SB")
        rptCell.Attributes.Add("colspan", 12)
        rptCell.InnerHtml = ""

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)

        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:16pt;font-family:DFKai-SB")
        rptCell.Attributes.Add("colspan", 12)
        rptCell.InnerHtml = Me.ViewState("PlanName")       '計劃名稱
        ' ---------
        rptTb = New HtmlTable
        rptTb.Attributes.Add("style", "BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
        rptTb.Attributes.Add("align", "left")
        rptTb.Attributes.Add("border", 0)
        print_content.Controls.Add(rptTb)
        '* ---班　　別： * 
        If dt4.Rows.Count > 0 Then
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
            rptCell.Attributes.Add("colspan", 12)
            rptCell.InnerHtml = "&nbsp班　　別： " & Convert.ToString(dt4.Rows(0)("ClassCName"))

            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
            rptCell.Attributes.Add("colspan", 12)
            rptCell.InnerHtml = "&nbsp訓練期間：" & Convert.ToDateTime(dt4.Rows(0)("STDate")).ToString("yyyy/MM/dd") & "～" & Convert.ToDateTime(dt4.Rows(0)("FTDate")).ToString("yyyy/MM/dd")
        Else
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
            rptCell.Attributes.Add("colspan", 12)
            rptCell.InnerHtml = "&nbsp班　　別："  ' 

            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
            rptCell.Attributes.Add("colspan", 12)
            rptCell.InnerHtml = "&nbsp訓練期間："
        End If
        '* ---列印日期 * 
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)

        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
        rptCell.Attributes.Add("colspan", 12)
        rptCell.InnerHtml = "&nbsp列印日期： " & Now.ToString("yyyy/MM/dd")
        '   
        Dim nl As HtmlGenericControl = New HtmlGenericControl
        print_content.Controls.Add(nl)
        nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;'>"
        '  - 依科目數目決定分頁  

        '  - 建立報表內容 start
        For i = 0 To dt3.Rows.Count - 1
            If i <= PageNum * PageSize - 1 And i > (PageNum - 1) * PageSize - 1 And i < (PageNum + 1) * PageSize - PageSize Then
                If (i <> 0 And (i Mod PageSize) = 0) Or (i = 0) Then
                    ' ----
                    rptTb = New HtmlTable
                    rptTb.Attributes.Add("style", "BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                    rptTb.Attributes.Add("align", "left")
                    rptTb.Attributes.Add("border", 0)
                    print_content.Controls.Add(rptTb)

                    rptRow = New HtmlTableRow
                    rptTb.Controls.Add(rptRow)
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "left")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("colspan", 12)
                    rptCell.InnerHtml = " "
                    ' --
                    rptTb = New HtmlTable
                    rptTb.Attributes.Add("align", "left")
                    rptTb.Attributes.Add("border", 1)
                    rptTb.Attributes.Add("style", "FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE:2px solid;BORDER-LEFT-STYLE:2px solid;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE:1px   solid")
                    rptTb.Attributes.Add("bordercolor", "black")

                    print_content.Controls.Add(rptTb)
                    '* --1 欄位名稱 --*   

                    rptRow = New HtmlTableRow
                    rptTb.Controls.Add(rptRow)

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.InnerHtml = "學號"

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.InnerHtml = "姓名"

                    If percent1 <> 0 And percent2 <> 0 And ResultTyp = 1 Then
                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("width", "70px")
                        rptCell.InnerHtml = "學科平均"

                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("width", "70px")
                        rptCell.InnerHtml = "術科平均"
                    End If

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "70px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = "總成績"

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "70px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = "總平均"

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "70px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = "操行"

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "70px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = "名次"

                    If dt1.Rows.Count > 0 Then
                        'Dim m As Integer = 0
                        For m As Integer = 0 To dt1.Rows.Count - 1
                            If m <= CourseNumSize * CoursePageNum - 1 And m > (CoursePageNum - 1) * CourseNumSize - 1 Then
                                dr1 = dt1.Rows(m)
                                rptCell = New HtmlTableCell
                                rptRow.Controls.Add(rptCell)
                                rptCell.Attributes.Add("align", "left")
                                rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                                rptCell.Attributes.Add("width", "100px")
                                rptCell.Attributes.Add("colspan", 1)
                                rptCell.InnerHtml = Convert.ToString(dr1("CourseName"))        '科目名稱
                            End If
                        Next
                    End If
                End If
                '* ---2 開始填入資料  --*   
                Dim dr3 As DataRow
                dr3 = dt3.Rows(i)
                rptRow = New HtmlTableRow
                rptTb.Controls.Add(rptRow)

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                rptCell.Attributes.Add("colspan", 1)
                rptCell.Attributes.Add("width", "80px")
                rptCell.InnerHtml = Convert.ToString(dr3("StudentID"))     '學號"

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                'rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                'word-wrap:break-word; overflow@hidden;
                rptCell.Attributes.Add("style", "word-wrap:break-word; overflow@hidden;font-size:12pt;font-family:DFKai-SB")
                rptCell.Attributes.Add("colspan", 1)
                rptCell.Attributes.Add("width", "80px")
                rptCell.InnerHtml = Convert.ToString(dr3("Name"))          '姓名"

                Dim dt As New DataTable
                Dim dr As DataRow
                dt = CreateGradeTable(sm.UserInfo.TPlanID, sm.UserInfo.DistID, OCIDValue1, Convert.ToString(dr3("socid")))

                ' --  
                'If dt.Rows.Count > 0 Then
                If dt IsNot Nothing Then
                    dr = dt.Rows(0)
                    ' 備註
                    If recordNote = "" Then
                        If Convert.ToString(dr("percent1")) = "0" And Convert.ToString(dr("percent2")) = "0" Then
                            If Convert.ToString(dr("ResultTyp")) = "1" Then
                                recordNote = " 備註:名次是由全部學員全部成績計算統計排名 "
                            Else
                                recordNote = " 備註:名次是由全部學員全部成績計算統計排名 <br/> "
                                recordNote += "總平均是由(各科成績與訓練時數加權/各科時數加總  " & Convert.ToString(Convert.ToInt16(dr("pHours")) + Convert.ToInt16(dr("sHours"))) & " 小時) 計算產生之結果。 "
                            End If
                        Else
                            If Convert.ToString(dr("ResultTyp")) = "1" Then  '成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
                                recordNote = " 備註: <br/>"
                                recordNote &= " 1.名次是由全部學員全部成績計算統計排名 <br/>"
                                recordNote &= " 2.總平均是由<br/>"
                                recordNote += "(學科全部成績加總/學科總數  " & Convert.ToString(dr("pClassCount")) & "  科) * 學科百分比 " & Convert.ToString(dr("percent1")) * 100 & " % <br/>"
                                recordNote += "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp加上<br/>"
                                recordNote += "(術科全部成績加總/術科總數  " & Convert.ToString(dr("sClassCount")) & "  科) * 術科百分比 " & Convert.ToString(dr("percent2")) * 100 & " % <br/>"
                                recordNote += "&nbsp&nbsp計算之結果。 "
                            Else
                                recordNote = " 備註: <br/>"
                                recordNote &= " 1.名次是由全部學員全部成績計算統計排名  <br/>"
                                recordNote &= " 2.總平均是由<br/>"
                                recordNote += "(學科成績與訓練時數加權 /學科時數 " & Convert.ToString(dr("pHours")) & " 小時) * 學科百分比 " & Convert.ToString(dr("percent1")) * 100 & " % <br/>"
                                recordNote += "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp加上<br/>"
                                recordNote += "(術科成績與訓練時數加權 /術科時數 " & Convert.ToString(dr("sHours")) & " 小時) * 術科百分比 " & Convert.ToString(dr("percent2")) * 100 & " % <br/>"
                                recordNote += "&nbsp&nbsp計算之結果。 "
                            End If
                        End If
                    End If
                    Me.ViewState("recordNote") = recordNote
                    ' ---
                    If percent1 <> 0 And percent2 <> 0 And ResultTyp = 1 Then
                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("width", "80px")

                        rptCell.InnerHtml = IIf(PrintDataTyp = "1", "&nbsp", Convert.ToString(dr("pAvg")))

                        'rptCell.InnerHtml = Convert.ToString(dr("pAvg")) '學科平均

                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("width", "80px")

                        rptCell.InnerHtml = IIf(PrintDataTyp = "1", "&nbsp", Convert.ToString(dr("sAvg")))

                        'rptCell.InnerHtml = Convert.ToString(dr("sAvg")) '術科平均
                    End If

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.Attributes.Add("colspan", 1)

                    If Convert.ToString(dr("ResultTyp")) = "1" Then
                        If Not dr("totalResults") Is Nothing Then
                            rptCell.InnerHtml = IIf(PrintDataTyp = "1", "&nbsp", Convert.ToString(dr("totalResults")))   '總成績"
                        Else
                            rptCell.InnerHtml = ""
                        End If
                    Else
                        If Not dr("totalResults2") Is Nothing Then
                            rptCell.InnerHtml = IIf(PrintDataTyp = "1", "&nbsp", Convert.ToString(dr("totalResults2")))
                        Else
                            rptCell.InnerHtml = ""
                        End If
                    End If
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = IIf(PrintDataTyp = "1", "&nbsp", IIf(IsDBNull(dr("totalAvg")), "", Convert.ToString(dr("totalAvg"))))  '總平均"
                    ' -將總平均逐個填入空的datatable
                    drAvg = dtAvg.NewRow
                    drAvg("socid") = Convert.ToString(dr3("socid"))
                    drAvg("totalAvg") = Convert.ToSingle(dr("totalAvg"))
                    dtAvg.Rows.Add(drAvg)
                Else
                    If percent1 <> 0 And percent2 <> 0 And ResultTyp = 1 Then
                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("width", "80px")
                        rptCell.InnerHtml = "" '學科平均

                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("width", "80px")
                        rptCell.InnerHtml = ""                                    '術科平均
                    End If
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = ""                                    '總成績"

                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = ""                                    '總平均"
                End If

                If dt4.Rows.Count > 0 Then
                    'Dim o As Integer = 0
                    Dim fff4 As String = String.Format("socid={0}", Convert.ToString(dr3("socid")))
                    If dt4.Select(fff4).Length > 0 Then
                        Dim dr4 As DataRow = dt4.Select(fff4)(0)
                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("width", "80px")
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = If(PrintDataTyp = "1", "&nbsp", Convert.ToString(dr4("conductpoint"))) '操行"
                    End If

                    'For o As Integer = 0 To dt4.Rows.Count - 1
                    '    If Convert.ToString(dt4.Rows(o)("socid")) = Convert.ToString(dr3("socid")) Then
                    '        rptCell = New HtmlTableCell
                    '        rptRow.Controls.Add(rptCell)
                    '        rptCell.Attributes.Add("align", "center")
                    '        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    '        rptCell.Attributes.Add("width", "80px")
                    '        rptCell.Attributes.Add("colspan", 1)
                    '        rptCell.InnerHtml = If(PrintDataTyp = "1", "&nbsp", Convert.ToString(dt4.Rows(o)("conductpoint")))
                    '    End If
                    'Next
                Else
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.Attributes.Add("colspan", 1)
                    rptCell.InnerHtml = ""  '操行"
                End If

                If dt5.Rows.Count > 0 Then
                    'Dim n As Integer = 0
                    Dim fff5 As String = String.Format("socid={0}", Convert.ToString(dr3("socid")))
                    If dt5.Select(fff5).Length > 0 Then
                        Dim dr5 As DataRow = dt5.Select(fff5)(0)
                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                        rptCell.Attributes.Add("align", "center")
                        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                        rptCell.Attributes.Add("width", "80px")
                        rptCell.Attributes.Add("colspan", 1)
                        If Convert.ToString(dr5("studstatus")) = "2" OrElse Convert.ToString(dr5("studstatus")) = "3" Then
                            rptCell.InnerHtml = ""
                        Else
                            '名次 
                            rptCell.InnerHtml = If(PrintDataTyp = "1", "&nbsp", Convert.ToString(dr5("rank")))
                        End If
                    End If

                    'For n As Integer = 0 To dt5.Rows.Count - 1
                    '    If Convert.ToString(dt5.Rows(n)("socid")) = Convert.ToString(dr3("socid")) Then
                    '        rptCell = New HtmlTableCell
                    '        rptRow.Controls.Add(rptCell)
                    '        rptCell.Attributes.Add("align", "center")
                    '        rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    '        rptCell.Attributes.Add("width", "80px")
                    '        rptCell.Attributes.Add("colspan", 1)
                    '        If Convert.ToString(dt5.Rows(n)("studstatus")) = "2" Or Convert.ToString(dt5.Rows(n)("studstatus")) = "3" Then
                    '            rptCell.InnerHtml = ""
                    '        Else
                    '            '名次 
                    '            rptCell.InnerHtml = If(PrintDataTyp = "1", "&nbsp", Convert.ToString(dt5.Rows(n)("rank")))
                    '        End If
                    '    End If
                    'Next
                Else
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("width", "80px")
                    rptCell.Attributes.Add("colspan", 1)
                    '名次 
                    rptCell.InnerHtml = ""
                End If

                'Dim dr4 As DataRow
                If dt1.Rows.Count > 0 Then                                      '依序取出學員各科目成績
                    'Dim m As Integer = 0
                    For m As Integer = 0 To dt1.Rows.Count - 1
                        If m <= CourseNumSize * CoursePageNum - 1 And m > (CoursePageNum - 1) * CourseNumSize - 1 Then
                            dr1 = dt1.Rows(m)
                            rptCell = New HtmlTableCell
                            rptRow.Controls.Add(rptCell)
                            rptCell.Attributes.Add("align", "left")
                            rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:12pt;font-family:DFKai-SB")
                            rptCell.Attributes.Add("width", "80px")
                            rptCell.Attributes.Add("colspan", 1)
                            If dt2.Select("(socid=" & Convert.ToString(dr3("socid")) & " or socid  is  null )  and CourID='" & Convert.ToString(dr1("courid")) & "'").Length = 0 Then
                                rptCell.InnerHtml = ""
                            Else
                                rptCell.InnerHtml = IIf(PrintDataTyp = "1", "&nbsp", Convert.ToString(dt2.Select("( socid= " & Convert.ToString(dr3("socid")) & " or socid  is  null )  and CourID='" & Convert.ToString(dr1("courid")) & "'").GetValue(0)("Results")))       '各科目成績
                            End If
                        End If
                    Next
                End If
            End If
        Next

        If dt1 IsNot Nothing Then dt1.Dispose()
        If dt2 IsNot Nothing Then dt2.Dispose()
        If dt3 IsNot Nothing Then dt3.Dispose()
        If dt4 IsNot Nothing Then dt4.Dispose()
        If dt2 IsNot Nothing Then dt2.Dispose()

    End Sub

    '建立成績檔案
    Private Function CreateGradeTable(ByVal TPlanID As String, ByVal DistID As String, ByVal OCIDValue1 As String, ByVal SOCID As String) As DataTable
        Dim Errmsg As String = ""
        'Dim i As Integer
        Dim dt As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        'Dim conn As New SqlConnection
        'Dim da As New SqlDataAdapter

        Dim pClassCount As Integer = 0              '學科科目總數
        Dim sClassCount As Integer = 0              '術科科目總數


        Call TIMS.OpenDbConn(objconn)

        Dim sql As String = "SELECT * FROM SYS_GLOBALVAR WHERE DistID=@DistID and TPlanID=@TPlanID"
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("DistID", SqlDbType.VarChar).Value = DistID
            .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID
            dt.Load(.ExecuteReader())
        End With

        Dim sql3 As String = ""
        sql3 &= " SELECT CourID,CourseName" & vbCrLf  '先取出目前所選取的科目  
        sql3 &= " ,case Classification1 when 1 then '學科'  when  2  then  '術科' end  classType" & vbCrLf  '先取出目前所選取的科目  
        sql3 &= " ,Classification1" & vbCrLf
        sql3 &= " FROM Course_CourseInfo" & vbCrLf  '先取出目前所選取的科目  
        sql3 &= " WHERE CourID in (" & ChooseClass & " )" & vbCrLf
        sql3 &= " Order By CourseName "
        Dim sCmd3 As New SqlCommand(sql3, objconn)
        With sCmd3
            dt3.Load(.ExecuteReader())
        End With

        If dt Is Nothing OrElse dt.Select("GVID='13'").Length = 0 Then
            'Errmsg += "尚未設定成績計算模式,請聯絡中心系統管理者" & vbCrLf
            Errmsg += "尚未設定成績計算模式,請聯絡分署系統管理者" & vbCrLf
            'Exit Sub
        Else
            CheckMode = dt.Select("GVID='13'")(0)("ItemVar1")
        End If

        If dt3.Rows.Count > 0 Then
            Dim n As Integer = 0
            For n = 0 To dt3.Rows.Count - 1
                If Convert.ToString(dt3.Rows(n)("Classification1")) = "1" Then
                    pClassCount = pClassCount + 1
                ElseIf Convert.ToString(dt3.Rows(n)("Classification1")) = "2" Then
                    sClassCount = sClassCount + 1
                End If
            Next
        End If
        'da.Dispose()
        dt.Dispose()
        dt3.Dispose()

        'Dim ChooseClass As String = Request("ChooseClass")  'Get_ChooseClass(OCIDValue1)
        If ChooseClass = "" Then
            lbmsg.Text = "查無資料"
            Return Nothing
        End If

        '建立資料    End
        'lbmsg.Text = "查無資料"

        Call TIMS.OpenDbConn(objconn)
        dt2 = getStudResult(2, SOCID)
        '    ----  
        '***************************************
        ' 【計算成績方式】
        '        <各科平均法>
        '=======================================
        '  (1)學、術科百分比(無)  (2) 學、術科百分比(有)
        '    ex: 總平均是由(學科全部成績加總/學科總數 4  科) * 學科百分比 20 %      
        '        加上      (術科全部成績加總/術科總數 11 科) * 術科百分比 80 %      
        '        計算產生之結果。
        '  ---------
        '        <訓練時數權重法> 
        '======================================= 
        '  (1)學、術科百分比(無)  
        '   總平均=(科)成績*(科)總時數 /(各科加總)總時數  
        '   總平均是由(各科成績與訓練時數加權/各科時數加總 __ 小時)計算產生之結果。 
        'ex:
        '總平均=
        '(學科成績與時數加權/時數加總 36 小時)* 學科百分比 40 %   
        '     +
        '(術科成績與時數加權/時數加總 88 小時)* 術科百分比 60 %        
        '/2 計算之結果。 

        '  ---------
        '  (2) 學、術科百分比(有)
        '    ex: 學科平均=(科)成績*(科)總時數 /(各科加總)總時數  
        '        總平均是由(學科成績與訓練時數加權/學科時數 __ 小時) * 學科百分比 20 %      
        '        加上      (術科成績與訓練時數加權/術科時數 __ 小時) * 術科百分比 20 %        
        '        計算產生之結果。 
        '***************************************
        Dim j As Integer = 0
        'Dim pClassCount As Integer = 0              '學科科目總數
        'Dim sClassCount As Integer = 0              '術科科目總數
        Dim pTotal_1 As Decimal = 0        '學科總成績---各科平均法
        Dim sTotal_1 As Decimal = 0        '術科總成績 
        Dim pTotal_2 As Decimal = 0        '學科總成績---訓練時數權重法
        Dim sTotal_2 As Decimal = 0        '術科總成績
        Dim pHours As Decimal = 0          '學科總時數
        Dim sHours As Decimal = 0          '術科總時數
        Dim totalResults As Decimal = 0    '總成績 
        Dim totalResults2 As Decimal = 0   '總成績(時數加權) 
        Dim totalAvg As Decimal = 0        '總平均 
        Dim pAvg As Decimal = 0            '學科總平均 
        Dim sAvg As Decimal = 0            '術科總平均 
        'Dim ResultTyp As Integer = GetPara(13, DistID, TPlanID)      '成績計算方式
        'Dim percent1 As Double = GetPara(17, DistID, TPlanID, 1)     '學科百分比
        'Dim percent2 As Double = GetPara(17, DistID, TPlanID, 2)     '術科百分比
        Dim s_ResultTyp As String = If(ViewState("ResultTyp") IsNot Nothing, Convert.ToString(ViewState("ResultTyp")), "")
        Dim ResultTyp As Double = GetPara(objconn, SET_PARMS1(13, sm, 1, s_ResultTyp)) '成績計算方式 ( 1, 各科平均法  2, 訓練時數權重法 )
        Dim percent1 As Double = GetPara(objconn, SET_PARMS1(17, sm, 1, s_ResultTyp)) '學科百分比
        Dim percent2 As Double = GetPara(objconn, SET_PARMS1(17, sm, 2, s_ResultTyp)) '術科百分比

        If dt2.Rows.Count > 0 Then
            For j = 0 To dt2.Rows.Count - 1
                If Convert.ToString(dt2.Rows(j)("Classification1")) = "1" Then        '學科
                    'pClassCount = pClassCount + 1
                    pHours = pHours + Convert.ToInt16(dt2.Rows(j)("hours"))
                    pTotal_2 = pTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
                    pTotal_1 = pTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))   '各科平均法： 1 '預設採用各科平均法(當未設定時)
                ElseIf Convert.ToString(dt2.Rows(j)("Classification1")) = "2" Then    '術科
                    'sClassCount = sClassCount + 1
                    sHours = sHours + Convert.ToInt16(dt2.Rows(j)("hours"))
                    sTotal_2 = sTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
                    sTotal_1 = sTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))
                End If
            Next

            If pClassCount = 0 Then
                pAvg = 0
            Else
                pAvg = pTotal_1 / pClassCount
            End If

            If sClassCount = 0 Then
                sAvg = 0
            Else
                sAvg = sTotal_1 / sClassCount
            End If

            totalResults = sTotal_1 + pTotal_1  '原始成績加總
            totalResults2 = sTotal_2 + pTotal_2 '成績加總(訓練時數權重法)

            Select Case ResultTyp
                Case 2        '訓練時數權重法： 2
                    '總平均= (學科成績與時數加權/時數加總 36 小時)* 學科百分比 40% + (術科成績與時數加權/時數加總 88 小時)* 術科百分比 60 %        
                    '/2 計算之結果。 
                    If percent1 = 0 And percent2 = 0 Then  '百分比未設定時
                        If pClassCount = 0 And sClassCount <> 0 And sHours <> 0 Then      '只有術科
                            totalAvg = (sTotal_2 / sHours)
                        ElseIf pClassCount <> 0 And sClassCount = 0 And pHours <> 0 Then  '只有學科
                            totalAvg = (pTotal_2 / pHours)
                        ElseIf pClassCount <> 0 And sClassCount <> 0 Then
                            If sHours = 0 And pHours <> 0 Then
                                totalAvg = (pTotal_2 / pHours)
                            ElseIf sHours <> 0 And pHours = 0 Then
                                totalAvg = (sTotal_2 / sHours)
                            ElseIf sHours <> 0 And pHours <> 0 Then
                                totalAvg = ((sTotal_2 / sHours) + (pTotal_2 / pHours)) / 2
                            End If
                        Else
                            totalAvg = 0
                        End If
                    Else
                        If pClassCount = 0 And sClassCount <> 0 And sHours <> 0 Then      '只有術科
                            totalAvg = (sTotal_2 / sHours) * percent2
                        ElseIf pClassCount <> 0 And sClassCount = 0 And pHours <> 0 Then  '只有學科
                            totalAvg = (pTotal_2 / pHours) * percent1
                        ElseIf pClassCount <> 0 And sClassCount <> 0 Then
                            If sHours = 0 And pHours <> 0 Then
                                totalAvg = (pTotal_2 / pHours) * percent1
                            ElseIf sHours <> 0 And pHours = 0 Then
                                totalAvg = (sTotal_2 / sHours) * percent2
                            ElseIf sHours <> 0 And pHours <> 0 Then
                                totalAvg = (pTotal_2 / pHours) * percent1 + (sTotal_2 / sHours) * percent2
                            Else
                                totalAvg = 0
                            End If
                        Else
                            totalAvg = 0
                        End If
                    End If
                Case Else     '各科平均法： 1 '預設採用各科平均法(當未設定時)
                    If percent1 = 0 And percent2 = 0 Then  '百分比未設定時
                        totalAvg = totalResults / (pClassCount + sClassCount)
                    Else
                        If pClassCount = 0 And sClassCount <> 0 Then      '只有術科
                            totalAvg = (sTotal_1 / sClassCount) * percent2
                        ElseIf pClassCount <> 0 And sClassCount = 0 Then  '只有學科
                            totalAvg = (pTotal_1 / pClassCount) * percent1
                        ElseIf pClassCount <> 0 And sClassCount <> 0 Then
                            totalAvg = (pTotal_1 / pClassCount) * percent1 + (sTotal_1 / sClassCount) * percent2
                        Else
                            totalAvg = 0
                        End If
                    End If
            End Select
            lbmsg.Text = ""

            Dim dtReult As New DataTable
            Dim dr As DataRow

            dtReult.Columns.Add(New DataColumn("SOCID"))
            dtReult.Columns.Add(New DataColumn("pClassCount"))
            dtReult.Columns.Add(New DataColumn("sClassCount"))
            dtReult.Columns.Add(New DataColumn("pTotal_1"))
            dtReult.Columns.Add(New DataColumn("sTotal_1"))
            dtReult.Columns.Add(New DataColumn("pTotal_2"))
            dtReult.Columns.Add(New DataColumn("sTotal_2"))
            dtReult.Columns.Add(New DataColumn("pHours"))
            dtReult.Columns.Add(New DataColumn("sHours"))
            dtReult.Columns.Add(New DataColumn("totalResults"))
            dtReult.Columns.Add(New DataColumn("totalResults2"))
            dtReult.Columns.Add(New DataColumn("totalAvg"))
            dtReult.Columns.Add(New DataColumn("pAvg"))
            dtReult.Columns.Add(New DataColumn("sAvg"))
            dtReult.Columns.Add(New DataColumn("ResultTyp"))
            dtReult.Columns.Add(New DataColumn("percent1"))
            dtReult.Columns.Add(New DataColumn("percent2"))

            dr = dtReult.NewRow
            dtReult.Rows.Add(dr)
            dr("SOCID") = SOCID
            dr("pClassCount") = pClassCount      '學科科目總數
            dr("sClassCount") = sClassCount      '術科科目總數
            dr("pTotal_1") = pTotal_1            '學科總成績---各科平均法
            dr("sTotal_1") = sTotal_1            '術科總成績 
            dr("pTotal_2") = pTotal_2            '學科總成績---訓練時數權重法 
            dr("sTotal_2") = sTotal_2            '術科總成績 
            dr("pHours") = pHours                '學科總時數
            dr("sHours") = sHours                '術科總時數
            dr("totalResults") = TIMS.ROUND(totalResults, 1)      '總成績 
            dr("totalResults2") = TIMS.ROUND(totalResults2, 1)     '總成績(時數加權) 
            dr("totalAvg") = TIMS.ROUND(totalAvg, 1)     '總平均 
            dr("pAvg") = TIMS.ROUND(pAvg, 1)             '學科總平均 
            dr("sAvg") = TIMS.ROUND(sAvg, 1)             '術科總平均 
            dr("ResultTyp") = ResultTyp   '成績計算方式 
            dr("percent1") = percent1     '學科百分比
            dr("percent2") = percent2     '術科百分比

            Return dtReult
        End If
        dt2.Dispose()

        Return Nothing
    End Function

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Private Sub bt_excel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles bt_excel.Click
        Dim fileName As String = "學員成績表.xls"
        If Request.Browser.Browser = "IE" Then
            fileName = Server.UrlPathEncode(fileName)
        End If
        Dim strContentDisposition As String = [String].Format("{0}; filename=""{1}""", "attachment", fileName)
        Response.AddHeader("Content-Disposition", strContentDisposition)
        Response.ContentType = "application/vnd.ms-excel"
        Dim sw As New System.IO.StringWriter
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        print_content.RenderControl(htw)
        Common.RespWrite(Me, sw.ToString().Replace("<div>", "").Replace("</div>", ""))
        Response.End()
    End Sub

    Private Sub tb_Page_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_Page.TextChanged
        If tb_Page.Text <> "" Then
            Dim PerDataNum As Integer = (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize))                                                  '資料依分頁共有幾頁來做循環
            Dim CoursePageNum As Integer = Math.Ceiling(Convert.ToInt16(tb_Page.Text) / (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize)))  '每頁可呈現之科目欄位數
            Dim PageIndexOf As Integer = Convert.ToInt16(tb_Page.Text) Mod PerDataNum   '呈現第幾筆資料
            If PageIndexOf = 0 Then PageIndexOf = PerDataNum

            Call CreateRpt(PageIndexOf, CoursePageNum)
        End If
    End Sub

    Private Sub bt_next_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles bt_next.Click
        If tb_PageTotal.Text <> "" Then
            If Convert.ToInt16(tb_Page.Text) = Convert.ToInt16(tb_PageTotal.Text) Then
                Common.MessageBox(Me.Page, "已是最後一頁!")
            Else
                tb_Page.Text = Convert.ToString(Convert.ToInt16(tb_Page.Text) + 1)
            End If
        End If
        Dim PerDataNum As Integer = (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize))                                                  '資料依分頁共有幾頁來做循環
        Dim CoursePageNum As Integer = Math.Ceiling(Convert.ToInt16(tb_Page.Text) / (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize)))  '每頁可呈現之科目欄位數
        Dim PageIndexOf As Integer = Convert.ToInt16(tb_Page.Text) Mod PerDataNum   '呈現第幾筆資料
        If PageIndexOf = 0 Then PageIndexOf = PerDataNum

        Call CreateRpt(PageIndexOf, CoursePageNum)
        If tb_Page.Text <> "" And tb_PageTotal.Text <> "" Then
            If Convert.ToInt16(tb_Page.Text) = Convert.ToInt16(tb_PageTotal.Text) Then
                Call printNote()
            End If
        End If
    End Sub

    Private Sub bt_end_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles bt_end.Click
        If tb_PageTotal.Text <> "" Then
            tb_Page.Text = tb_PageTotal.Text
        End If
        Dim PerDataNum As Integer = (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize))                                                 '資料依分頁共有幾頁來做循環
        Dim CoursePageNum As Integer = Math.Ceiling(Convert.ToInt16(tb_Page.Text) / (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize)))  '每頁可呈現之科目欄位數
        Dim PageIndexOf As Integer = Convert.ToInt16(tb_Page.Text) Mod PerDataNum   '呈現第幾筆資料
        If PageIndexOf = 0 Then PageIndexOf = PerDataNum

        Call CreateRpt(PageIndexOf, CoursePageNum)
        If tb_Page.Text <> "" And tb_PageTotal.Text <> "" Then
            If Convert.ToInt16(tb_Page.Text) = Convert.ToInt16(tb_PageTotal.Text) Then
                Call printNote()
            End If
        End If
    End Sub

    Private Sub bt_first_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles bt_first.Click
        If tb_PageTotal.Text <> "" Then
            tb_Page.Text = "1"
        End If
        Dim PerDataNum As Integer = (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize))                                                  '資料依分頁共有幾頁來做循環
        Dim CoursePageNum As Integer = Math.Ceiling(Convert.ToInt16(tb_Page.Text) / (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize))) '每頁可呈現之科目欄位數
        Dim PageIndexOf As Integer = Convert.ToInt16(tb_Page.Text) Mod PerDataNum   '呈現第幾筆資料
        If PageIndexOf = 0 Then PageIndexOf = PerDataNum

        Call CreateRpt(PageIndexOf, CoursePageNum)
    End Sub

    Private Sub bt_pre_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles bt_pre.Click
        If tb_PageTotal.Text <> "" Then
            If Convert.ToInt16(tb_Page.Text) = "1" Then
                Common.MessageBox(Me.Page, "已是第一頁!")
            Else
                tb_Page.Text = Convert.ToString(Convert.ToInt16(tb_Page.Text) - 1)
            End If
        End If
        Dim PerDataNum As Integer = (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize))                                                  '資料依分頁共有幾頁來做循環
        Dim CoursePageNum As Integer = Math.Ceiling(Convert.ToInt16(tb_Page.Text) / (Convert.ToInt16(tb_PageTotal.Text) / Math.Ceiling(GetCourseCount() / CourseNumSize))) '每頁可呈現之科目欄位數
        Dim PageIndexOf As Integer = Convert.ToInt16(tb_Page.Text) Mod PerDataNum   '呈現第幾筆資料
        If PageIndexOf = 0 Then PageIndexOf = PerDataNum

        Call CreateRpt(PageIndexOf, CoursePageNum)
    End Sub

    Private Sub lb_print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_print.Click
        Call PrintResult()
    End Sub

End Class