Partial Class SD_03_007
    Inherits AuthBasePage
    '    Dim blnCanAdds As Boolean = False '新增,Dim blnCanMod As Boolean = False '修改,Dim blnCanDel As Boolean = False '刪除,Dim blnCanSech As Boolean = False '查詢,Dim blnCanPrnt As Boolean = False '列印,

    'Dim s_Coloum As String()
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        'https://jira.turbotech.com.tw/browse/TIMSC-127
        RedStar.Visible = If(TIMS.Cst_TPlanIDCanNoUseOCID.IndexOf(sm.UserInfo.TPlanID) > -1, False, True)

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Button1.Attributes("onclick") = "javascript:return search();" '班級必選

            If RedStar.Visible Then
                If sm.UserInfo.TPlanID = TIMS.Cst_TPlanID06 Then
                    msg.Text = "[自辦職前]未選擇班級時依登入年度、轄區計畫、訓練機構匯出資料!!"
                End If
            End If

            Call Create_chkSort(chkSort) '建立匯出欄位資料

            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
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

        chkSort.Attributes("onclick") = "SelectAll('chkSort','chkSortHidden');"
        'Button1.Attributes("onclick") = "javascript:return search();"
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    '匯出Excel button
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Creattable() '匯出Excel
    End Sub

    ''' <summary> 建立匯出欄位資料(column) </summary>
    ''' <returns></returns>
    Function Get_ColumnStr1() As String
        'DIM SORTARY() AS STRING
        Dim str_SortAry As String = ""
        str_SortAry = ""
        str_SortAry &= "全部"
        str_SortAry &= ",DISTNAME" '轄區" 'DISTNAME
        str_SortAry &= ",CLASSCNAME" '班別名稱"  '班別名稱'CLASSCNAME
        str_SortAry &= ",CYCLTYPE" '期別" 'CYCLTYPE

        str_SortAry &= ",STUDENTID" '學號" 'STUDENTID
        str_SortAry &= ",SNAME" '中文姓名" 'SNAME
        str_SortAry &= ",ENGNAME" '英文姓名" 'ENGNAME
        str_SortAry &= ",IDNO_MK" 'IDNO 身分證號碼" 'IDNO
        str_SortAry &= ",SEX" '性別" 'SEX

        str_SortAry &= ",PASSPORTNAME" '身分別" 'PASSPORTNAME
        str_SortAry &= ",CHINAORNOT" '非本國人身分別" 'CHINAORNOT
        str_SortAry &= ",NATIONALITY" '原屬國籍" 'NATIONALITY
        str_SortAry &= ",PPNO" '護照或工作證號" 'PPNO
        str_SortAry &= ",BIRTHDAY_MK" ',BIRTHDAY" '出生日期" 'BIRTHDAY
        str_SortAry &= ",YEARSOLD" '年齡" 'YEARSOLD
        'str_SortAry &= ",MARITALSTATUS" '婚姻狀況" 'MARITALSTATUS

        str_SortAry &= ",DEGREEID" '最高學歷" 'DEGREEID
        str_SortAry &= ",SCHOOL" '學校名稱" 'SCHOOL
        str_SortAry &= ",DEPARTMENT" '科系" 'DEPARTMENT
        str_SortAry &= ",GRADID" '畢業狀況" 'GRADID

        str_SortAry &= ",MILITARYID" '兵役" 'MILITARYID
        str_SortAry &= ",SERVICEID" '軍種" 'SERVICEID
        'STR_SORTARY &= ",兵役職務" 'MILITARYAPPOINTMENT
        'STR_SORTARY &= ",兵役階級" 'MILITARYRANK
        str_SortAry &= ",SERVICEORG" '服役單位名稱" 'SERVICEORG
        'STR_SORTARY &= ",主管階級姓名" 'CHIEFRANKNAME
        str_SortAry &= ",SSERVICEDATE" '服役起日期" 'SSERVICEDATE
        str_SortAry &= ",FSERVICEDATE" '服役迄日期" 'FSERVICEDATE
        'STR_SORTARY &= ",服役單位地址郵遞區號" 'ZIPCODE4
        'STR_SORTARY &= ",服役單位地址" 'SERVICEADDRESS
        str_SortAry &= ",SERVICEPHONE" '服役單位電話" 'SERVICEPHONE

        str_SortAry &= ",PHONED" '聯絡電話_日" 'PHONED
        str_SortAry &= ",PHONEN" '聯絡電話_夜" 'PHONEN
        str_SortAry &= ",CELLPHONE" '行動電話" 'CELLPHONE
        str_SortAry &= ",ZIPCODE1" '通訊地址郵遞區號" 'ZIPCODE1
        str_SortAry &= ",ADDRESS" '通訊地址" 'ADDRESS
        str_SortAry &= ",ZIPCODE2" '戶籍地址郵遞區號" 'ZIPCODE2
        str_SortAry &= ",HOUSEHOLDADDRESS" '戶籍地址" 'HOUSEHOLDADDRESS

        str_SortAry &= ",EMAIL" 'E_MAIL" 'EMAIL
        str_SortAry &= ",IDENTITYID" '參訓身分別" 'IDENTITYID
        str_SortAry &= ",MIDENTITYID" '主要參訓身分別" 'MIDENTITYID
        'STR_SORTARY &= ",津貼類別" 'SUBSIDYID
        str_SortAry &= ",OPENDATE" '開訓日期" 'OPENDATE
        str_SortAry &= ",CLOSEDATE" '結訓日期" 'CLOSEDATE
        str_SortAry &= ",ENTERDATE" '報到日期" 'ENTERDATE

        str_SortAry &= ",HANDTYPEID" '障礙類別" 'HANDTYPEID2/HANDTYPEID
        str_SortAry &= ",HANDLEVELID" '障礙等級" 'HANDLEVELID2/HANDLEVELID
        'str_SortAry &= ",EMERGENCYCONTACT" '緊急通知人姓名" 'EMERGENCYCONTACT
        'str_SortAry &= ",EMERGENCYRELATION" '緊急通知人關係" 'EMERGENCYRELATION
        'str_SortAry &= ",EMERGENCYPHONE" '緊急通知人電話" 'EMERGENCYPHONE
        'str_SortAry &= ",ZIPCODE3" '緊急通知人地址郵遞區號" 'ZIPCODE3
        'str_SortAry &= ",EMERGENCYADDRESS" '緊急通知人地址" 'EMERGENCYADDRESS

        'STR_SORTARY &= ",交通方式" 'TRAFFIC
        str_SortAry &= ",SHOWDETAIL" '是否提供基本資料查詢" 'SHOWDETAIL
        'STR_SORTARY &= ",報名階段" 'LEVELNO
        str_SortAry &= ",ENTERCHANNEL" '報名管道" 'ENTERCHANNEL

        str_SortAry &= ",BUDNAME" '預算別" 'BUDNAME
        str_SortAry &= ",ISAGREE" '個資法意願" 'ISAGREE
        'STR_SORTARY &= ",自費/公費" 'PMODE
        'STR_SORTARY &= ",國內親屬資料_姓名" 'FORENAME
        'STR_SORTARY &= ",國內親屬資料_稱謂" 'FORETITLE
        'STR_SORTARY &= ",國內親屬資料_性別" 'FORESEX
        'STR_SORTARY &= ",國內親屬資料_生日" 'FOREBIRTH
        'STR_SORTARY &= ",國內親屬資料_身分證號碼" 'FOREIDNO
        'STR_SORTARY &= ",國內親屬資料_郵遞區號" 'FOREZIP
        'STR_SORTARY &= ",國內親屬資料_地址" 'FOREADDR
        str_SortAry &= ",NATIVEN" '原住民民族別" 'NATIVEN

        str_SortAry &= ",REJECTTDATE1" '離訓日期" 'REJECTTDATE1
        str_SortAry &= ",REJECTTDATE2" '退訓日期" 'REJECTTDATE2
        'STR_SORTARY &= ",是否就業"                 '(依照承辦人需求,將此欄位拿掉，BY:20180919)
        'STR_SORTARY &= ",在職者身分"  '在職者補助身分'WORKSUPPIDENT
        str_SortAry &= ",ACTNO" '投保證號" 'ACTNO
        str_SortAry &= ",ACTNAME" '投保單位名稱" 'ACTNAME
        Return str_SortAry
    End Function

    ''' <summary> 建立匯出欄位資料(title) </summary>
    ''' <param name="chkobj"></param>
    Public Shared Sub Create_chkSort(ByRef chkobj As CheckBoxList)
        Dim str_SortAry As String = ""
        str_SortAry = ""
        str_SortAry &= "全部"
        str_SortAry &= ",轄區" 'DISTNAME
        str_SortAry &= ",班別名稱"  '班別名稱'CLASSCNAME
        str_SortAry &= ",期別" 'CYCLTYPE

        str_SortAry &= ",學號" 'StudentID
        str_SortAry &= ",中文姓名" 'SName
        str_SortAry &= ",英文姓名" 'EngName
        str_SortAry &= ",身分證號碼" 'IDNO
        str_SortAry &= ",性別" 'Sex

        str_SortAry &= ",身分別" 'PassPortName
        str_SortAry &= ",非本國人身分別" 'ChinaOrNot
        str_SortAry &= ",原屬國籍" 'Nationality
        str_SortAry &= ",護照或工作證號" 'PPNO
        str_SortAry &= ",出生日期" 'Birthday
        str_SortAry &= ",年齡" 'YEARSOLD
        'str_SortAry &= ",婚姻狀況" 'MaritalStatus

        str_SortAry &= ",最高學歷" 'DegreeID
        str_SortAry &= ",學校名稱" 'School
        str_SortAry &= ",科系" 'Department
        str_SortAry &= ",畢業狀況" 'GradID

        str_SortAry &= ",兵役" 'MilitaryID
        str_SortAry &= ",軍種" 'ServiceID
        'str_SortAry &= ",兵役職務" 'MilitaryAppointment
        'str_SortAry &= ",兵役階級" 'MilitaryRank
        str_SortAry &= ",服役單位名稱" 'ServiceOrg
        'str_SortAry &= ",主管階級姓名" 'ChiefRankName
        str_SortAry &= ",服役起日期" 'SServiceDate
        str_SortAry &= ",服役迄日期" 'FServiceDate
        'str_SortAry &= ",服役單位地址郵遞區號" 'ZipCode4
        'str_SortAry &= ",服役單位地址" 'ServiceAddress
        str_SortAry &= ",服役單位電話" 'ServicePhone

        str_SortAry &= ",聯絡電話_日" 'PhoneD
        str_SortAry &= ",聯絡電話_夜" 'PhoneN
        str_SortAry &= ",行動電話" 'CellPhone
        str_SortAry &= ",通訊地址郵遞區號" 'ZipCode1
        str_SortAry &= ",通訊地址" 'Address
        str_SortAry &= ",戶籍地址郵遞區號" 'ZipCode2
        str_SortAry &= ",戶籍地址" 'HouseholdAddress

        str_SortAry &= ",E_Mail" 'Email
        str_SortAry &= ",參訓身分別" 'IdentityID
        str_SortAry &= ",主要參訓身分別" 'MIdentityID
        'str_SortAry &= ",津貼類別" 'SubsidyID
        str_SortAry &= ",開訓日期" 'OpenDate
        str_SortAry &= ",結訓日期" 'CloseDate
        str_SortAry &= ",報到日期" 'EnterDate

        str_SortAry &= ",障礙類別" 'HandTypeID2/HandTypeID
        str_SortAry &= ",障礙等級" 'HandLevelID2/HandLevelID
        'str_SortAry &= ",緊急通知人姓名" 'EmergencyContact
        'str_SortAry &= ",緊急通知人關係" 'EmergencyRelation
        'str_SortAry &= ",緊急通知人電話" 'EmergencyPhone
        'str_SortAry &= ",緊急通知人地址郵遞區號" 'ZipCode3
        'str_SortAry &= ",緊急通知人地址" 'EmergencyAddress

        'str_SortAry &= ",交通方式" 'Traffic
        str_SortAry &= ",是否提供基本資料查詢" 'ShowDetail
        'str_SortAry &= ",報名階段" 'LevelNo
        str_SortAry &= ",報名管道" 'EnterChannel

        str_SortAry &= ",預算別" 'BudName
        str_SortAry &= ",個資法意願" 'IsAgree
        'str_SortAry &= ",自費/公費" 'PMode
        'str_SortAry &= ",國內親屬資料_姓名" 'ForeName
        'str_SortAry &= ",國內親屬資料_稱謂" 'ForeTitle
        'str_SortAry &= ",國內親屬資料_性別" 'ForeSex
        'str_SortAry &= ",國內親屬資料_生日" 'ForeBirth
        'str_SortAry &= ",國內親屬資料_身分證號碼" 'ForeIDNO
        'str_SortAry &= ",國內親屬資料_郵遞區號" 'ForeZip
        'str_SortAry &= ",國內親屬資料_地址" 'ForeAddr
        str_SortAry &= ",原住民民族別" 'NativeN

        str_SortAry &= ",離訓日期" 'RejectTDate1
        str_SortAry &= ",退訓日期" 'RejectTDate2
        'str_SortAry &= ",是否就業"  '(依照承辦人需求,將此欄位拿掉，by:20180919)
        'str_SortAry &= ",在職者身分"  '在職者補助身分'WorkSuppIdent
        str_SortAry &= ",投保證號" 'ACTNO
        str_SortAry &= ",投保單位名稱" 'ACTNAME

        Dim SortAry() As String = str_SortAry.Split(",")
        '75+2
        With chkobj
            .Items.Clear()
            For i As Integer = 0 To SortAry.Length - 1
                .Items.Add(New ListItem(SortAry(i), CStr(i)))
            Next
        End With
        'Sort.Items.Clear()
    End Sub

    ''' <summary>取得資料欄位值</summary>
    ''' <param name="IdenDt1"></param>
    ''' <param name="dr"></param>
    ''' <param name="s_Coloum"></param>
    ''' <param name="iChk_a"></param>
    ''' <returns></returns>
    Function Get_dr_value2(ByRef IdenDt1 As DataTable, ByRef dr As DataRow, ByRef s_Coloum As String(), ByRef iChk_a As Integer) As String
        Dim rst As String = ""
        Select Case s_Coloum(iChk_a)
            Case "IDENTITYID" '參訓身分別
                rst = TIMS.Get_IdentityName(dr("IdentityID").ToString, IdenDt1, ",", "，")
            Case "MIDENTITYID" '主要參訓身分別
                rst = TIMS.Get_IdentityName(dr("MIdentityID").ToString, IdenDt1, ",", "，")
            Case "HANDTYPEID" '障礙類別 
                '障礙類別2 
                rst = If(Convert.ToString(dr("HandTypeID2")) <> "", TIMS.Get_HandTypeName2(dr("HandTypeID2")), TIMS.ClearSQM(dr("HandTypeID")))
            Case "HANDLEVELID" '障礙等級 
                '障礙等級2 
                rst = If(Convert.ToString(dr("HandLevelID2")) <> "", TIMS.ClearSQM(dr("HandLevelID2Name")), TIMS.ClearSQM(dr("HandTypeID")))
            Case "EMERGENCYPHONE"
                Dim vEmergencyPhone As String = TIMS.ClearSQM(dr("EmergencyPhone"))
                If vEmergencyPhone <> "" AndAlso vEmergencyPhone.Contains(",") Then vEmergencyPhone = Replace(vEmergencyPhone, ",", "、")
                rst = "'" & vEmergencyPhone '緊急通知人電話
            Case Else
                rst = TIMS.ClearSQM(dr(s_Coloum(iChk_a)))
        End Select
        Return rst
    End Function

    '匯出Excel
    Sub Creattable()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        Dim dt As DataTable = Nothing
        'Dim ExportStr As String = ""

        Dim MIdentityID As String = ""
        Dim IdentityID As String = ""

        '讀取Class_ClassInfo、Class_StudentsOfClass、Stud_StudentInfo、Stud_SubData、Key_Degree、Key_Military、Key_Subsidy、Key_HandicatType、
        'Key_HandicatLevel、Key_JoblessWeek、Key_Budget、Key_Native

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT C2.StudentID" & vbCrLf
        sql &= " ,C2.SOCID" & vbCrLf
        sql &= " ,S1.Name SName" & vbCrLf
        sql &= " ,S1.EngName" & vbCrLf
        sql &= " ,S1.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(S1.IDNO) IDNO_MK" & vbCrLf
        sql &= " ,dbo.DECODE6(S1.Sex,'M','男','F','女','') Sex" & vbCrLf
        sql &= " ,dbo.DECODE6(S1.PassPortNO,1,'本國',2,'外籍','') PassPortName" & vbCrLf
        sql &= " ,case S1.ChinaOrNot when 1 then '是' else '否' end ChinaOrNot" & vbCrLf
        sql &= " ,S1.Nationality" & vbCrLf
        sql &= " ,S1.PPNO" & vbCrLf
        sql &= " ,FORMAT(S1.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK2(S1.Birthday) BIRTHDAY_MK" & vbCrLf
        sql &= " ,CASE S1.MaritalStatus WHEN 1 THEN '已婚' WHEN 2 THEN '未婚' END MaritalStatus" & vbCrLf
        sql &= " ,K1.Name DegreeID" & vbCrLf
        sql &= " ,S2.School" & vbCrLf
        sql &= " ,S2.Department" & vbCrLf
        sql &= " ,K2.Name MilitaryID" & vbCrLf
        sql &= " ,S2.ServiceID" & vbCrLf
        sql &= " ,S2.ServiceOrg" & vbCrLf
        sql &= " ,CASE S1.GraduateStatus WHEN '01' THEN '畢業' WHEN '02' THEN '肄業' ELSE '在學中' END GradID" & vbCrLf
        sql &= " ,FORMAT(S2.SServiceDate,'yyyy/MM/dd') SServiceDate" & vbCrLf
        sql &= " ,FORMAT(S2.FServiceDate,'yyyy/MM/dd') FServiceDate" & vbCrLf
        sql &= " ,S2.ServicePhone" & vbCrLf
        sql &= " ,S2.PhoneD" & vbCrLf
        sql &= " ,S2.PhoneN" & vbCrLf
        sql &= " ,S2.CellPhone" & vbCrLf
        sql &= " ,dbo.FN_GZIP33(S2.ZipCode1,S2.ZipCode1_6W) ZipCode1" & vbCrLf
        sql &= " ,I2.ZipName+S2.Address Address" & vbCrLf
        sql &= " ,dbo.FN_GZIP33(S2.ZipCode2,S2.ZipCode2_6W) ZipCode2" & vbCrLf
        sql &= " ,I3.ZipName+S2.HouseholdAddress HouseholdAddress" & vbCrLf
        sql &= " ,S2.Email" & vbCrLf
        sql &= " ,C2.IdentityID" & vbCrLf
        sql &= " ,C2.MIdentityID" & vbCrLf
        sql &= " ,K3.Name SubsidyID" & vbCrLf
        sql &= " ,FORMAT(C2.OpenDate,'yyyy/MM/dd') OpenDate" & vbCrLf
        sql &= " ,FORMAT(C2.CloseDate,'yyyy/MM/dd') CloseDate" & vbCrLf
        sql &= " ,FORMAT(C2.EnterDate,'yyyy/MM/dd') EnterDate" & vbCrLf
        sql &= " ,FORMAT(C2.RejectTDate1,'yyyy/MM/dd') RejectTDate1" & vbCrLf
        sql &= " ,FORMAT(C2.RejectTDate2,'yyyy/MM/dd') RejectTDate2" & vbCrLf

        sql &= " ,K4.Name HandTypeID" & vbCrLf
        sql &= " ,K5.Name HandLevelID" & vbCrLf
        sql &= " ,S2.HandTypeID2" & vbCrLf
        sql &= " ,S2.HandLevelID2" & vbCrLf

        sql &= " ,S2.EmergencyContact" & vbCrLf
        sql &= " ,S2.EmergencyRelation" & vbCrLf
        sql &= " ,S2.EmergencyPhone" & vbCrLf
        sql &= " ,dbo.FN_GZIP33(S2.ZipCode3,S2.ZipCode3_6W) ZipCode3" & vbCrLf
        sql &= " ,ISNULL(I4.ZipName,'')+ISNULL(S2.EmergencyAddress,'') EmergencyAddress" & vbCrLf
        sql &= " ,S2.PriorWorkOrg1 ,S2.Title1 ,S2.SOfficeYM1 ,S2.FOfficeYM1" & vbCrLf
        sql &= " ,S2.PriorWorkOrg2 ,S2.Title2 ,S2.SOfficeYM2 ,S2.FOfficeYM2" & vbCrLf
        sql &= " ,S2.PriorWorkPay ,S1.RealJobless" & vbCrLf
        'sql &= " /*　,K6.Name AS JoblessID" & vbCrLf
        'sql &= " ,CASE S2.Traffic WHEN 1 THEN '住宿' WHEN 2 THEN '通勤' END AS Traffic*/" & vbCrLf
        sql &= " ,CASE S2.ShowDetail WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END AS ShowDetail" & vbCrLf
        'sql &= " /*,C2.LevelNo*/" & vbCrLf
        sql &= " ,CASE C2.EnterChannel WHEN 1 THEN '網路' WHEN 2 THEN '現場' WHEN 3 THEN '通訊' WHEN 4 THEN '推介' END AS EnterChannel" & vbCrLf
        'sql &= " /*,CASE C2.TRNDMode WHEN 1 THEN '職訓券' WHEN 2 THEN '學習券' WHEN 3 THEN '推介券' END AS TRNDMode*/" & vbCrLf
        sql &= " ,K7.BudName" & vbCrLf
        sql &= " ,CASE S1.IsAgree WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END AS IsAgree" & vbCrLf
        'sql &= " /*,CASE C2.TRNDType WHEN 1 THEN '甲式' WHEN 2 THEN '乙式' END AS TRNDType" & vbCrLf
        'sql &= " ,CASE C2.PMode WHEN 1 THEN '公費' WHEN 2 THEN '自費' END AS PMode" & vbCrLf
        'sql &= " ,S2.ForeName ,S2.ForeTitle ,S2.ForeBirth ,S2.ForeIDNO" & vbCrLf
        'sql &= " ,CASE S2.ForeSex WHEN 'M' THEN '男' WHEN 'F' THEN '女' END AS ForeSex" & vbCrLf
        'sql &= " ,I5.ZipName + S2.ForeAddr AS ForeAddr" & vbCrLf
        'sql &= " ,CASE ISNULL(C2.WorkSuppIdent,'N') WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END AS WorkSuppIdent" & vbCrLf
        'sql &= " ,CASE C2.IsOnJob WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END AS IsOnJob" & vbCrLf
        'sql &= " */" & vbCrLf
        sql &= " ,K8.Name NativeN" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(C1.CLASSCNAME,C1.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,C1.CLASSCNAME" & vbCrLf
        sql &= " ,C1.CYCLTYPE" & vbCrLf

        sql &= " ,K5b.Name HandLevelID2Name" & vbCrLf
        sql &= " ,C1.RID" & vbCrLf
        sql &= " ,ip.YEARS" & vbCrLf
        sql &= " ,ip.TPlanID" & vbCrLf
        sql &= " ,ip.PlanID" & vbCrLf
        sql &= " ,ip.DistID" & vbCrLf
        sql &= " ,ip.DISTNAME" & vbCrLf
        sql &= " ,C1.OCID" & vbCrLf
        sql &= " ,DATEDIFF(YEAR,S1.BIRTHDAY,C1.STDATE) YEARSOLD" & vbCrLf
        sql &= " ,SP.ACTNO" & vbCrLf
        sql &= " ,SP.ACTNAME" & vbCrLf
        'sql &= " /* 「轄區」、「期別」、「年齡」、「投保證號」、「投保單位名稱」 */" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO C1" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip ON ip.PlanID = C1.PlanID" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP rr ON rr.RID = C1.RID" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS C2 ON C2.OCID=C1.OCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO S1 ON C2.SID = S1.SID" & vbCrLf
        sql &= " JOIN STUD_SUBDATA S2 ON S1.SID = S2.SID" & vbCrLf
        sql &= " JOIN KEY_DEGREE K1 ON K1.DegreeID = S1.DegreeID" & vbCrLf
        sql &= " LEFT JOIN STUD_SERVICEPLACE SP ON SP.SOCID=C2.SOCID" & vbCrLf
        sql &= " LEFT JOIN KEY_MILITARY K2 ON S1.MilitaryID = K2.MilitaryID" & vbCrLf
        'sql &= " /*LEFT JOIN VIEW_ZIPNAME I1 ON I1.ZipCode = S2.ZipCode4*/" & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME I2 ON I2.ZipCode = S2.ZipCode1" & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME I3 ON I3.ZipCode = S2.ZipCode2" & vbCrLf
        sql &= " LEFT JOIN KEY_SUBSIDY K3 ON C2.SubsidyID = K3.SubsidyID" & vbCrLf
        sql &= " LEFT JOIN KEY_HANDICATTYPE K4 ON S2.HandTypeID = K4.HandTypeID" & vbCrLf
        sql &= " LEFT JOIN KEY_HANDICATLEVEL K5 ON S2.HandLevelID = K5.HandLevelID" & vbCrLf
        sql &= " LEFT JOIN KEY_HANDICATLEVEL2 K5b ON S2.HandLevelID2 = K5b.HandLevelID2" & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME I4 ON I4.ZipCode = S2.ZipCode3" & vbCrLf
        'sql &= " /*LEFT JOIN KEY_JOBLESSWEEK K6 ON S1.JoblessID = K6.JoblessID*/" & vbCrLf
        sql &= " LEFT JOIN KEY_BUDGET K7 ON C2.BudgetID = K7.BudID" & vbCrLf
        'sql &= " /*LEFT JOIN VIEW_ZIPNAME I5 ON I5.ZipCode = S2.ForeZip*/" & vbCrLf
        sql &= " LEFT JOIN KEY_NATIVE K8 ON C2.Native = K8.KNID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND C1.ISSUCCESS = 'Y'" & vbCrLf
        sql &= " AND C1.NOTOPEN = 'N'" & vbCrLf
        'sql &= " AND ip.Years = '2020'" & vbCrLf
        'sql &= " AND ip.TPlanID = '28'" & vbCrLf
        sql &= " AND rr.RID = @RID" & vbCrLf '機構業務
        sql &= " AND ip.Years = @Years" & vbCrLf '年度
        sql &= " AND ip.TPlanID = @TPlanID" & vbCrLf '大計畫
        If OCIDValue1.Value <> "" Then sql &= " AND C1.OCID= @OCIDV1" & vbCrLf  'OCIDValue1

        Dim parms As Hashtable = New Hashtable()
        parms.Clear()
        parms.Add("RID", RIDValue.Value) '機構業務
        parms.Add("Years", sm.UserInfo.Years) '年度
        parms.Add("TPlanID", sm.UserInfo.TPlanID) '大計畫
        If OCIDValue1.Value <> "" Then parms.Add("OCIDV1", Val(OCIDValue1.Value)) 'OCIDValue1

        Select Case Convert.ToString(sm.UserInfo.LID) '署(局)
            Case "0"
                'sql &= " AND rr.RID = @RID " & vbCrLf '機構業務
                'sql &= " AND ip.Years = @Years " & vbCrLf '年度
                'sql &= " AND ip.TPlanID = @TPlanID " & vbCrLf '大計畫
                'parms.Add("RID", RIDValue.Value)
                'parms.Add("Years", sm.UserInfo.Years)
                'parms.Add("TPlanID", sm.UserInfo.TPlanID)
            Case Else
                sql &= " AND ip.PlanID = @PlanID" & vbCrLf '小計畫
                sql &= " AND ip.DistID = @DistID" & vbCrLf '轄區
                'sql &= " AND rr.RID = @RID " & vbCrLf '機構業務
                'sql &= " AND ip.Years = @Years " & vbCrLf '年度
                'sql &= " AND ip.TPlanID = @TPlanID " & vbCrLf '大計畫
                parms.Add("PlanID", sm.UserInfo.PlanID)
                parms.Add("DistID", sm.UserInfo.DistID)
                'parms.Add("RID", RIDValue.Value)
                'parms.Add("Years", sm.UserInfo.Years)
                'parms.Add("TPlanID", sm.UserInfo.TPlanID)
        End Select
        'sql &= " ORDER BY C2.StudentID " & vbCrLf

        'dt = DbAccess.GetDataTable(sql, objconn, parms)
        Dim flag_error As Boolean = True '預設為錯誤 ! 查詢正確時為false 
        'Dim dt As DataTable = Nothing
        Try
            dt = DbAccess.GetDataTable(sql, objconn, parms)
            flag_error = False
        Catch ex As Exception
            Dim cst_fun_page_name As String = "##SD_03_007.aspx, "
            Dim slogMsg1 As String = ""
            slogMsg1 &= String.Concat(cst_fun_page_name, "sql: ", sql) & vbCrLf
            slogMsg1 &= String.Concat(cst_fun_page_name, "parms: ", TIMS.GetMyValue3(parms)) & vbCrLf
            'Call TIMS.SendMailTest(slogMsg1)
            Dim strErrmsg As String = ""
            strErrmsg &= "ex.Message: " & vbCrLf & ex.Message & vbCrLf
            strErrmsg &= "ex.ToString: " & vbCrLf & ex.ToString & vbCrLf
            strErrmsg &= "slogMsg1: " & vbCrLf & slogMsg1 & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.SendMailTest(strErrmsg)
        End Try
        If flag_error Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg13)
            Exit Sub
        End If
        If dt Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        dt.DefaultView.Sort = "StudentID"
        'If OCIDValue1.Value <> "" Then
        '    '因為資料庫加入OCID的條件 ，就會變得超慢，所以改為程式過濾
        '    dt.DefaultView.RowFilter = "OCID=" & OCIDValue1.Value
        'End If
        dt = TIMS.dv2dt(dt.DefaultView)
        If dt Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        ExportX1(dt)
    End Sub

    Sub ExportX1(ByRef dt As DataTable)
        If dt Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim sFileName1 As String = String.Concat("學員資料", TIMS.GetDateNo2())

        Dim sMemo As String = ""
        sMemo = ""
        sMemo &= "&ACT=" & sFileName1 & vbCrLf
        'sMemo &= String.Format("&parms=({0})", TIMS.GET_PARMSVAL(parms)) & vbCrLf
        sMemo &= "&OCIDValue1=" & OCIDValue1.Value & vbCrLf
        sMemo &= "&COUNT=" & dt.Rows.Count & vbCrLf
        TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip1, OCIDValue1.Value, sMemo)

        Dim ExportStr As String = ""

        Dim sqlstr As String = ""
        sqlstr = " SELECT * FROM KEY_IDENTITY ORDER BY IDENTITYID "
        Dim IdenDt1 As DataTable = DbAccess.GetDataTable(sqlstr, objconn)

        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")
        'Const cst_Split As String = "<,>" '分隔號

        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" border=""1"">")
        'strHTML &= ("<table>")

        '匯出欄位全部 
        Dim v_chkSort As String = TIMS.GetCblValue(chkSort)
        Dim flag_OutColAll As Boolean = If(v_chkSort = "", True, False)

        '建立輸出文字
        ExportStr = ""
        For iChk_a As Integer = 0 To chkSort.Items.Count - 1
            If flag_OutColAll Then '匯出欄位全部
                If iChk_a <> 0 Then
                    Dim TXT_chkSort As String = chkSort.Items.Item(iChk_a).Text
                    ExportStr &= String.Concat(If(ExportStr <> "", ",", ""), TXT_chkSort)
                End If
            Else
                If chkSort.Items.Item(iChk_a).Selected AndAlso iChk_a <> 0 Then
                    Dim TXT_chkSort As String = chkSort.Items.Item(iChk_a).Text
                    ExportStr &= String.Concat(If(ExportStr <> "", ",", ""), TXT_chkSort)
                End If
            End If
        Next
        strHTML &= TIMS.Get_TABLETR(ExportStr)

        Dim s_Coloum As String() = Get_ColumnStr1().Split(",")
        ExportStr = "" ' vbCrLf
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        For Each dr As DataRow In dt.Rows
            ExportStr = ""
            If flag_OutColAll Then
                For iChk_a As Integer = 0 To chkSort.Items.Count - 1
                    If iChk_a <> 0 Then
                        Dim TXT_chkSort As String = Get_dr_value2(IdenDt1, dr, s_Coloum, iChk_a)
                        ExportStr &= String.Concat(If(ExportStr <> "", TIMS.cst_SplitB1, ""), TXT_chkSort)
                    End If
                Next
            Else
                For iChk_j As Integer = 0 To chkSort.Items.Count - 1
                    If chkSort.Items.Item(iChk_j).Selected AndAlso iChk_j <> 0 Then
                        Dim TXT_chkSort As String = Get_dr_value2(IdenDt1, dr, s_Coloum, iChk_j)
                        ExportStr &= String.Concat(If(ExportStr <> "", TIMS.cst_SplitB1, ""), TXT_chkSort)
                    End If
                Next
            End If
            strHTML &= TIMS.Get_TABLETR(ExportStr, True, TIMS.cst_SplitB1)
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") '  Response.End()
    End Sub

    '判斷機構是否只有一個班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    '判斷機構是否只有一個班級
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class