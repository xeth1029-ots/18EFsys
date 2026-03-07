Public Class TC_03_004
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents HyperLink1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents File1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents dtgAddresses1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dtgAddresses2 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dtgAddresses4 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dtgAddresses3 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Btn_XlsImport As System.Web.UI.WebControls.Button

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region


    Dim Key_TrainType As DataTable
    Dim ID_GovClassCast As DataTable
    Dim Key_HourRan As DataTable
    Dim Key_ZipCode As DataTable

    
    Dim dt1, dt2, dt3, dt4 As DataTable

    'Const Cst_PlanID = 10
    'Const Cst_SeqNO = 10
    'Const Cst_PlanYear = 10
    'Const Cst_TPlanID = 10
    'Const Cst_CapSex = 12  
    'Const Cst_CapMilitary = 13
    Const Cst_ComIDNO = 0           '訓練機構統一編號 
    Const Cst_ClassName = 1         '班別名稱*  
    Const Cst_GCID = 2              '經費分類代碼 

    Const Cst_TMID = 3              '訓練業別代碼* 
    Const Cst_PlanEMail = 4         '線上報名Email 
    Const Cst_PlanCause = 5         '目標-緣由*
    Const Cst_PurScience = 6        '目標-學科* 
    Const Cst_PurTech = 7           '目標-技能* 
    Const Cst_PurMoral = 8          '目標-品德* 
    Const Cst_CapDegree = 9         '受訓資格-學歷代碼* 
    Const Cst_CapAge1 = 10          '受訓資格-年齡起* 
    Const Cst_CapAge2 = 11          '受訓資格-年齡迄* 
    Const Cst_CapOther1 = 12        '受訓資格-其它一
    Const Cst_CapOther2 = 13        '受訓資格-其它二
    Const Cst_CapOther3 = 14        '受訓資格-其它三
    Const Cst_TMScience = 15        '訓練方式* 

    Const Cst_FirstSort = 16        '優先排序* 
    Const Cst_PointYN = 17          '課程種類
    Const Cst_IsBusiness = 18       '企業包班
    Const Cst_EnterpriseName = 19   '企業包班公司名稱 

    Const Cst_TNum = 20             '訓練人數* 
    Const Cst_THours = 21           '訓練時數* 
    Const Cst_STDate = 22           '訓練起日*  
    Const Cst_FDDate = 23           '訓練迄日* 
    Const Cst_CyclType = 24         '期別(二碼)* 
    Const Cst_ClassCount = 25       '班數* 
    Const Cst_TaddressZip = 26      '上課地址(郵遞區號3碼)* 
    Const Cst_TaddressZip2w = 27    '上課地址(郵遞區號2碼)* 
    Const Cst_TAddress = 28         '上課地址* 
    Const Cst_CredPoint = 29        '學分數 
    Const Cst_SciPlaceID = 30       '學科場地* 
    Const Cst_TechPlaceID = 31      '術科場地* 
    Const Cst_ConNum = 32           '容納人數* 
    Const Cst_ContactName = 33      '聯絡人 
    Const Cst_ContactPhone = 34     '聯絡人電話 
    Const Cst_ContactFax = 35       '聯絡人傳真 
    Const Cst_ContactEmail = 36     '聯絡人Email 
    Const Cst_ClassCate = 37        '訓練職能*  
    Const Cst_EnterSupplyStyle = 38 '報名繳費方式 
    Const Cst_Note = 39             '訓練費用編列說明(備註) 

    ' Const Cst_TPeriod = 40          '上課時段* 
    Const Cst_TrainDemain = 40      '訓練需求調查 
    Const Cst_TeacherDesc = 41      '師資遴選辦法說明 
    Const Cst_CapAll = 42          '學員資格*  
    Const Cst_RecDesc = 43          '1#反應評估 
    Const Cst_LearnDesc = 44        '2#學習評估 
    Const Cst_ActDesc = 45          '3#行為評估 
    Const Cst_ResultDesc = 46       '4#成果評估 
    Const Cst_OtherDesc = 47        '5#其它機制 
    Const Cst_Recruit = 48          '招訓方式 
    Const Cst_Inspire = 49          '學員激勵辦法* 

    Const cst_filedNum = 50
    Const cst_必須填寫 = "必須填寫"

    '課程大綱
    Const Cst_STrainDate = 2 '日期 
    Const Cst_PName = 3 '授課時間 
    Const Cst_PHour = 4  '時數 
    Const Cst_PCont = 5 '課程進度／內容 
    Const Cst_Classification1 = 6 '學／術科 
    Const Cst_PTID = 7 '上課地點 
    Const Cst_TechID = 8 '任課教師 

    Const Cst_CostID = 2 '項目
    Const Cst_OPrice = 3 '單價
    Const Cst_Itemage = 4 '計價數量

    Const Cst_Weeks = 2 '星期x
    Const Cst_Times = 3 '時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

    End Sub

    Private Sub Btn_XlsImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_XlsImport.Click
        Const Cst_FileSavePath = "~/TC/01/Temp/"
        Const Cst_Filetype = "xls"

        Dim MyFile As System.IO.File
        Dim MyFileName As String
        Dim MyFileType As String
        Dim flag As String
        If File1.Value <> "" Then
            '檢查檔案格式與大小-----------------------------------------------------Start
            If File1.PostedFile.ContentLength = 0 Then
                Turbo.Common.MessageBox(Me, "檔案位置錯誤!")
                Exit Sub
            Else
                '取出檔案名稱
                MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then

                    Turbo.Common.MessageBox(Me, "檔案類型錯誤!")
                    Exit Sub
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If LCase(MyFileType) = LCase(Cst_Filetype) Then
                        flag = ","
                    Else
                        Turbo.Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!")
                        Exit Sub
                    End If
                End If
            End If
            '檢查檔案格式與大小-----------------------------------------------------End

            '上傳檔案
            File1.PostedFile.SaveAs(Server.MapPath(Cst_FileSavePath & MyFileName))

            Dim Reason As String                '儲存錯誤的原因

            dt1 = TIMS.GetDataTable_XlsFileSheet( _
                            Server.MapPath(Cst_FileSavePath & MyFileName).ToString, _
                            "班級申請", Reason, "訓練機構統一編號")

            dt2 = TIMS.GetDataTable_XlsFileSheet( _
                            Server.MapPath(Cst_FileSavePath & MyFileName).ToString, _
                            "課程大綱", Reason, "訓練機構統一編號")

            dt3 = TIMS.GetDataTable_XlsFileSheet( _
                            Server.MapPath(Cst_FileSavePath & MyFileName).ToString, _
                            "上課時間", Reason, "訓練機構統一編號")

            dt4 = TIMS.GetDataTable_XlsFileSheet( _
                            Server.MapPath(Cst_FileSavePath & MyFileName).ToString, _
                            "訓練費用", Reason, "訓練機構統一編號")

            MyFile.Delete(Server.MapPath(Cst_FileSavePath & MyFileName)) '刪除檔案

            'dtgAddresses1.DataSource = dt1
            'dtgAddresses1.DataBind()
            'dtgAddresses2.DataSource = dt2
            'dtgAddresses2.DataBind()
            'dtgAddresses3.DataSource = dt3
            'dtgAddresses3.DataBind()
            'dtgAddresses4.DataSource = dt4
            'dtgAddresses4.DataBind()
            'Exit Sub

            If Reason <> "" Then
                Turbo.Common.MessageBox(Me, Reason)
                Exit Sub
            End If

            'xls 方式 讀取寫入資料庫
            If dt1.Rows.Count > 0 Then '有資料
                '將檔案讀出放入記憶體

                Dim RowIndex As Integer = 1
                Dim OneRow As String
                Dim colArray As Array

                '取出資料庫的所有欄位---------------------------------------------------Start
                Dim sql As String
                Dim dr As DataRow
                Dim da As SqlDataAdapter
                Dim trans As SqlTransaction
                Dim conn As SqlConnection = DbAccess.GetConnection

                Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
                Dim drWrong As DataRow

                '建立錯誤資料格式Table----------------Start
                dtWrong.Columns.Add(New DataColumn("Index"))
                dtWrong.Columns.Add(New DataColumn("ClassName"))
                dtWrong.Columns.Add(New DataColumn("ComIDNO"))
                dtWrong.Columns.Add(New DataColumn("Reason"))
                '建立錯誤資料格式Table----------------End

                '取出所有鍵值當判斷-----------------------------------Start
                sql = " SELECT * FROM Key_TrainType where Parent IN (SELECT TMID FROM Key_TrainType where BusID='G')"
                Key_TrainType = DbAccess.GetDataTable(sql)
                sql = " SELECT * FROM ID_GovClassCast "
                ID_GovClassCast = DbAccess.GetDataTable(sql)
                sql = " SELECT * FROM Key_HourRan"
                Key_HourRan = DbAccess.GetDataTable(sql)
                sql = "SELECT * FROM ID_Zip"
                Key_ZipCode = DbAccess.GetDataTable(sql)
                '取出所有鍵值當判斷-----------------------------------End

                For i As Integer = 0 To dt1.Rows.Count - 1
                    Reason = ""
                    colArray = dt1.Rows(i).ItemArray

                    Reason += CheckImportData(colArray) '檢查及置放
                    '通過檢查，開始輸入資料---------------------Start

                    If Reason = "" Then
                        'colArray = ChangeImportDate(colArray)
                        If Not Save_Plan_PlanInfo(colArray, Reason) Then
                            '錯誤資料，填入錯誤資料表
                            drWrong = dtWrong.NewRow
                            dtWrong.Rows.Add(drWrong)
                            drWrong("Index") = RowIndex
                            If colArray.Length > 3 Then
                                drWrong("ClassName") = colArray(Cst_ClassName).ToString
                                drWrong("ComIDNO") = colArray(Cst_ComIDNO).ToString
                                drWrong("Reason") = Reason
                            End If
                        End If
                    Else
                        '錯誤資料，填入錯誤資料表
                        drWrong = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)
                        drWrong("Index") = RowIndex
                        If colArray.Length > 3 Then
                            drWrong("ClassName") = colArray(Cst_ClassName).ToString
                            drWrong("ComIDNO") = colArray(Cst_ComIDNO).ToString
                            drWrong("Reason") = Reason
                        End If
                    End If

                    RowIndex += 1
                Next

                '判斷匯出資料是否有誤
                Dim explain, explain2 As String
                explain = ""
                explain += "班級申請：匯入資料共" & dt1.Rows.Count & "筆" & vbCrLf
                explain += "成功：" & (dt1.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
                explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

                explain2 = ""
                explain2 += "班級申請：匯入資料共" & dt1.Rows.Count & "筆\n"
                explain2 += "成功：" & (dt1.Rows.Count - dtWrong.Rows.Count) & "筆\n"
                explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

                '開始判別欄位存入-------------------------------------------------------End
                If dtWrong.Rows.Count = 0 Then
                    Turbo.Common.MessageBox(Me, explain)
                Else
                    Session("MyWrongTable") = dtWrong
                    Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視失敗原因?')){window.open('TC_03_004_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
                End If 'dtWrong.Rows.Count
            End If 'If dt1.Rows.Count > 0 Then '有資料
        End If 'If File1.Value <> "" Then

    End Sub

    Function CheckImportData(ByVal colArray As Array)
        'Const cst_filedNum = 51
        'Const cst_必須填寫 = "必須填寫"

        Dim Reason As String
        Dim sql As String
        Dim dr As DataRow

        If colArray.Length <> cst_filedNum Then
            'Reason += "欄位數量不正確(應該為" & cst_filedNum & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
        Else
            Dim PlanIDValue As String = sm.UserInfo.PlanID
            Dim ComIDNO As String = colArray(Cst_ComIDNO).ToString '訓練機構統一編號            																																																																					                       
            Dim ClassName As String = colArray(Cst_ClassName).ToString  '班別名稱*       																																																																					                   
            Dim GCID As String = colArray(Cst_GCID).ToString  '經費分類代碼              																																																																					                   

            Dim TMID As String = colArray(Cst_TMID).ToString  '訓練業別代碼*             																																																																					                   
            Dim PlanEMail As String = colArray(Cst_PlanEMail).ToString  '線上報名Email   																																																																					                   
            Dim PlanCause As String = colArray(Cst_PlanCause).ToString  '目標-緣由*      																																																																					                   
            Dim PurScience As String = colArray(Cst_PurScience).ToString  '目標-學科*    																																																																					                   
            Dim PurTech As String = colArray(Cst_PurTech).ToString  '目標-技能*          																																																																					                   
            Dim PurMoral As String = colArray(Cst_PurMoral).ToString  '目標-品德*        																																																																					                   
            Dim CapDegree As String = colArray(Cst_CapDegree).ToString  '受訓資格-學歷代?																																																																				X*                 
            Dim CapAge1 As String = colArray(Cst_CapAge1).ToString  '受訓資格-年齡起*    																																																																					                   
            Dim CapAge2 As String = colArray(Cst_CapAge2).ToString  '受訓資格-年齡迄*    																																																																					                   
            Dim CapOther1 As String = colArray(Cst_CapOther1).ToString  '受訓資格-其它一 																																																																					                   
            Dim CapOther2 As String = colArray(Cst_CapOther2).ToString  '受訓資格-其它二 																																																																					                   
            Dim CapOther3 As String = colArray(Cst_CapOther3).ToString  '受訓資格-其它三 																																																																					                   
            Dim TMScience As String = colArray(Cst_TMScience).ToString  '訓練方式*       																																																																					                   

            Dim FirstSort As String = colArray(Cst_FirstSort).ToString  '優先排序*       																																																																					                   
            Dim PointYN As String = colArray(Cst_PointYN).ToString  '課程種類            																																																																					                   
            Dim IsBusiness As String = colArray(Cst_IsBusiness).ToString  '企業包班      																																																																					                   
            Dim EnterpriseName As String = colArray(Cst_EnterpriseName).ToString  '企業包																																																																					班公司名稱         

            Dim TNum As String = colArray(Cst_TNum).ToString  '訓練人數*                 																																																																					                   
            Dim THours As String = colArray(Cst_THours).ToString  '訓練時數*             																																																																					                   
            Dim STDate As String = colArray(Cst_STDate).ToString  '訓練起日*             																																																																					                   
            Dim FDDate As String = colArray(Cst_FDDate).ToString  '訓練迄日*             																																																																					                   
            Dim CyclType As String = colArray(Cst_CyclType).ToString  '期別(二碼)*       																																																																					                   
            Dim ClassCount As String = colArray(Cst_ClassCount).ToString  '班數*         																																																																					                   
            Dim TaddressZip As String = colArray(Cst_TaddressZip).ToString  '上課地址(郵?																																																																				摯牉?碼)*         
            Dim TaddressZip2w As String = colArray(Cst_TaddressZip2w).ToString  '上課地址																																																																					(郵遞區號2碼)*     
            Dim TAddress As String = colArray(Cst_TAddress).ToString  '上課地址*         																																																																					                   
            Dim CredPoint As String = colArray(Cst_CredPoint).ToString  '學分數          																																																																					                   
            Dim SciPlaceID As String = colArray(Cst_SciPlaceID).ToString  '學科場地*     																																																																					                   
            Dim TechPlaceID As String = colArray(Cst_TechPlaceID).ToString  '術科場地*   																																																																					                   
            Dim ConNum As String = colArray(Cst_ConNum).ToString  '容納人數*             																																																																					                   
            Dim ContactName As String = colArray(Cst_ContactName).ToString  '聯絡人      																																																																					                   
            Dim ContactPhone As String = colArray(Cst_ContactPhone).ToString  '聯絡人電話																																																																					                   
            Dim ContactFax As String = colArray(Cst_ContactFax).ToString  '聯絡人傳真    																																																																					                   
            Dim ContactEmail As String = colArray(Cst_ContactEmail).ToString  '聯絡人Emai																																																																					l                  
            Dim ClassCate As String = colArray(Cst_ClassCate).ToString  '訓練職能*       																																																																					                   
            Dim EnterSupplyStyle As String = colArray(Cst_EnterSupplyStyle).ToString '報?																																																																				W繳費方式          
            Dim Note As String = colArray(Cst_Note).ToString  '訓練費用編列說明(備註)    																																																																					                   

            'Dim TPeriod As String = colArray(Cst_TPeriod).ToString  '上課時段*           																																																																					                   
            Dim TrainDemain As String = colArray(Cst_TrainDemain).ToString  '訓練需求調查																																																																					                   
            Dim TeacherDesc As String = colArray(Cst_TeacherDesc).ToString  '師資遴選辦法																																																																					說明               
            Dim CapAll As String = colArray(Cst_CapAll).ToString  '學員資格*             																																																																					                   
            Dim RecDesc As String = colArray(Cst_RecDesc).ToString  '1#反應評估          																																																																					                   
            Dim LearnDesc As String = colArray(Cst_LearnDesc).ToString  '2#學習評估      																																																																					                   
            Dim ActDesc As String = colArray(Cst_ActDesc).ToString  '3#行為評估          																																																																					                   
            Dim ResultDesc As String = colArray(Cst_ResultDesc).ToString  '4#成果評估    																																																																					                   
            Dim OtherDesc As String = colArray(Cst_OtherDesc).ToString  '5#其它機制      																																																																					                   
            Dim Recruit As String = colArray(Cst_Recruit).ToString  '招訓方式            																																																																					                   
            Dim Inspire As String = colArray(Cst_Inspire).ToString  '學員激勵辦法*      

            If Trim(ComIDNO) = "" Then
                Reason += "必須填寫統一編號<Br>"
            Else
                If ComIDNO.Length <> 8 Then
                    Reason += "廠商統一編號必須為8碼<BR>"
                End If
                If Not IsNumeric(ComIDNO) Then
                    Reason += "廠商統一編號必須為數字<BR>"
                End If

                If TIMS.Get_RIDforOrgID(TIMS.Get_OrgIDforComIDNO(ComIDNO), PlanIDValue) = Nothing Then
                    Reason += " 此機構尚未加入此計劃<Br>"
                End If

                sql = "SELECT * FROM Org_OrgInfo WHERE ComIDNO='" & ComIDNO & "'"
                dr = DbAccess.GetOneRow(sql)
                If dr Is Nothing Then
                    Reason += "找不到相關訓練機構，請先新增訓練機構<BR>"
                End If
            End If

            If Trim(ClassName) = "" Then
                Reason += "必須填寫班別名稱<Br>"
            Else
                If ClassName.Length > 100 Then
                    Reason += "班別名稱長度超過，請修改<BR>"
                End If
            End If

            If Trim(GCID) <> "" Then
                Dim MyKey1 As String = TIMS.Get_GCIDCode(GCID)
                If MyKey1 Is Nothing Then
                    Reason += "經費分類代碼* 有錯，不符合鍵詞<BR>"
                Else
                    If ID_GovClassCast.Select("GCID='" & MyKey1 & "'").Length = 0 Then
                        Reason += "經費分類代碼* 有錯，不符合鍵詞<BR>"
                    End If
                End If
                colArray(Cst_GCID) = MyKey1
            Else
                Reason += "必須填寫經費分類代碼* <BR>"
            End If

            If Trim(TMID) <> "" Then
                Dim MyKey1 As String = TIMS.Get_jobTMID(TMID)
                If MyKey1 Is Nothing Then
                    Reason += "訓練業別代碼* 有錯，不符合鍵詞<BR>"
                Else
                    If Key_TrainType.Select("TMID='" & MyKey1 & "'").Length = 0 Then
                        Reason += "訓練業別代碼* 有錯，不符合鍵詞<BR>"
                    End If
                End If
                colArray(Cst_TMID) = MyKey1
            Else
                Reason += "必須填寫訓練業別代碼* <BR>"
            End If

            If Trim(PlanEMail) = "" Then
                Reason += "必須填寫線上報名Email <BR>"
            End If

            If Trim(PlanCause) <> "" Then
                If PlanCause.Length > 300 Then
                    Reason += "目標-緣由*內容長度超過，請修改<BR>"
                End If
            Else
                Reason += "必須填寫目標-緣由*<Br>"
            End If

            If Trim(PurScience) <> "" Then
                If PurScience.Length > 300 Then
                    Reason += "目標-學科*內容長度超過，請修改<BR>"
                End If
            Else
                Reason += "必須填寫目標-學科*<Br>"
            End If

            If Trim(PurTech) <> "" Then
                If PurTech.Length > 300 Then
                    Reason += "目標-技能*內容長度超過，請修改<BR>"
                End If
            Else
                Reason += "必須填寫目標-技能*<Br>"
            End If

            If Trim(PurMoral) <> "" Then
                If PurMoral.Length > 300 Then
                    Reason += "目標-品德*內容長度超過，請修改<BR>"
                End If
            Else
                Reason += "必須填寫目標-品德*<Br>"
            End If

            If TaddressZip <> "" Then
                If IsNumeric(TaddressZip) = False Then
                    Reason += "上課地址郵遞區號3碼必須要是數字<BR>"
                End If
                If Key_ZipCode.Select("ZipCode='" & TaddressZip & "'").Length = 0 Then
                    Reason += "上課地址郵遞區號3碼有錯，不符合鍵詞<BR>"
                End If
            End If

            If TaddressZip2w = "" Then
                Reason += "必須填寫上課地址郵遞區號2碼<BR>"
            Else
                If Len(TaddressZip2w) = 2 Then
                    If TaddressZip2w < "00" Or TaddressZip2w > "99" Then
                        Reason += "上課地址郵遞區號2碼填寫有誤 <BR>"
                    End If
                ElseIf Len(TaddressZip2w) = 1 Then
                    If TaddressZip2w < "0" Or TaddressZip2w > "9" Then
                        Reason += "上課地址郵遞區號2碼寫有誤 <BR>"
                    End If
                Else
                    Reason += "上課地址郵遞區號2碼填寫有誤 <BR>"
                End If
            End If
            'If Trim(TPeriod) <> "" Then
            '    Dim MyKey1 As String = TPeriod
            '    If MyKey1.Length = 1 Then
            '        MyKey1 = "0" & MyKey1
            '    End If
            '    If Key_HourRan.Select("HRID='" & MyKey1 & "'").Length = 0 Then
            '        Reason += "上課時段*代碼 有錯，不符合鍵詞<BR>"
            '    End If
            '    colArray(Cst_TPeriod) = MyKey1
            'Else
            '    Reason += "必須填寫 上課時段*代碼 <BR>"
            'End If

            If STDate <> "" Then
                If Not IsDate(STDate) Then
                    Reason += "訓練起日「日期」必須為正確的日期格式<BR>"
                End If
            Else
                Reason += "必須填寫訓練起日「日期」*<Br>"
            End If

            If FDDate <> "" Then
                If Not IsDate(FDDate) Then
                    Reason += "訓練迄日「日期」必須為正確的日期格式<BR>"
                End If
            Else
                Reason += "必須填寫訓練迄日「日期」*<Br>"
            End If

            If Trim(CapAll) <> "" Then
                If CapAll.Length > 200 Then
                    Reason += "學員資格*內容長度超過，請修改<BR>"
                End If
            Else
                Reason += "必須填寫 學員資格*<Br>"
            End If

            If Trim(Inspire) <> "" Then
                If Inspire.Length > 300 Then
                    Reason += "學員激勵辦法*內容長度超過，請修改<BR>"
                End If
            Else
                Reason += "必須填寫學員激勵辦法*<Br>"
            End If


            Dim str_Rule As String
            str_Rule = "[訓練機構統一編號]='" & ComIDNO & "'"
            '課程大綱
            For Each drw As DataRow In dt2.Select(str_Rule, Nothing, DataViewRowState.CurrentRows)
                If drw(Cst_ClassName).ToString = colArray(Cst_ClassName).ToString Then

                    If drw(Cst_STrainDate).ToString <> "" Then
                        If Not IsDate(drw(Cst_STrainDate).ToString) Then
                            Reason += "課程大綱「日期」必須為正確的日期格式<BR>"
                        End If
                    Else
                        Reason += "必須填寫課程大綱「日期」*<Br>"
                    End If

                    If drw(Cst_PName).ToString = "" Then
                        Reason += "必須填寫課程大綱「授課時間」*<Br>"
                    End If

                    '20090318--加入課程大綱時數須檢核為整數數字。
                    If drw(Cst_PHour).ToString = "" Then
                        Reason += "必須填寫課程大綱「時數」*<Br>"
                    Else
                        If IsNumeric(drw(Cst_PHour)) = False Then
                            Reason += "課程大綱「時數」*必須為數字<Br>"
                        Else
                            If CInt(drw(Cst_PHour)) <> drw(Cst_PHour) Then
                                Reason += "課程大綱「時數」*必須為整數<Br>"
                            End If
                        End If
                    End If

                    If drw(Cst_Classification1).ToString = "" Then
                        Reason += "必須填寫課程大綱「學／術科」*<Br>"
                    Else
                        If Not IsNumeric(drw(Cst_Classification1).ToString) Then
                            Reason += "課程大綱「學／術科」*，超出鍵詞範圍<Br>"
                        Else
                            Select Case CInt(drw(Cst_Classification1).ToString)
                                Case 1, 2
                                Case Else
                                    Reason += "課程大綱「學／術科」*，超出鍵詞範圍<Br>"
                            End Select
                        End If
                    End If

                    If drw(Cst_PTID).tostring <> "" Then
                        If TIMS.Get_PTID(drw(Cst_PTID).tostring, ComIDNO) = Nothing Then
                            Reason += "課程大綱「上課地點」*，超出鍵詞範圍<Br>"
                        End If
                    End If

                    If drw(Cst_TechID).tostring <> "" Then
                        If IsNumeric(drw(Cst_TechID)) Then
                            If TIMS.Get_TeacherName(drw(Cst_TechID)) = Nothing Then
                                Reason += "[" & drw(Cst_TechID).tostring & "]課程大綱「任課教師」*，超出鍵詞範圍<Br>"
                            End If
                        Else
                            Reason += "[" & drw(Cst_TechID).tostring & "]課程大綱「任課教師」*，超出鍵詞範圍<Br>"
                        End If
                    End If
                End If
            Next

            '上課時間
            For Each drw As DataRow In dt3.Select(str_Rule, Nothing, DataViewRowState.CurrentRows)
                If drw(Cst_ClassName).ToString = colArray(Cst_ClassName).ToString Then
                    Select Case drw(Cst_Weeks)
                        Case "1", "2", "3", "4", "5", "6", "7"
                        Case Else
                            Reason += "上課時間「星期」*，超出鍵詞範圍<Br>"
                    End Select
                End If
            Next

            '訓練費用
            Dim CostIDArray As New ArrayList
            For Each drw As DataRow In dt4.Select(str_Rule, Nothing, DataViewRowState.CurrentRows)
                If drw(Cst_ClassName).ToString = colArray(Cst_ClassName).ToString Then

                    Dim MYKEY As String = drw(Cst_CostID).toString
                    If MYKEY = "" Then
                        Reason += "必須填寫訓練費用「項目」*<Br>"
                    Else
                        '=========== 驗証匯入檔案時不要有相同的訓練費用「項目」 Start =============
                        Dim Flag As Boolean = True
                        For i As Integer = 0 To CostIDArray.Count - 1
                            If CostIDArray(i) = MYKEY Then
                                Reason += "檔案中有相同的訓練費用「項目」<BR>"
                                Flag = False
                            End If
                        Next
                        If Flag Then CostIDArray.Add(MYKEY)
                        '=========== 驗証匯入檔案時不要有相同的訓練費用「項目」 -End- =============
                    End If
                    If drw(Cst_OPrice).tostring = "" Then
                        Reason += "必須填寫訓練費用「單價」*<Br>"
                    Else
                        If Not IsNumeric(drw(Cst_OPrice).tostring) Then
                            Reason += "訓練費用「單價」*，必須為數字格式<Br>"
                        End If
                    End If
                    If drw(Cst_Itemage).tostring = "" Then
                        Reason += "必須填寫訓練費用「計價數量」*<Br>"
                    Else
                        If Not IsNumeric(drw(Cst_OPrice).tostring) Then
                            Reason += "訓練費用「計價數量」*，必須為數字格式<Br>"
                        End If
                    End If
                End If
            Next
        End If
        Return Reason
    End Function

    Function Save_Plan_PlanInfo(ByVal colArray As Array, ByRef Errmsg As String) As Boolean
        Save_Plan_PlanInfo = False
        Dim sql As String
        Dim da As SqlDataAdapter
        Dim dt As DataTable
        Dim dr As DataRow
        Dim Trans As SqlTransaction
        Dim conn As SqlConnection



        conn = DbAccess.GetConnection

        Try
            'TC_03_003.aspx
            Trans = DbAccess.BeginTrans(conn)

            Dim ComIDNOValue As String
            Dim int_SeqNO As Integer
            Dim RIDValue As String
            ComIDNOValue = colArray(Cst_ComIDNO).ToString
            int_SeqNO = GetMaxSeqNum(Trans, ComIDNOValue)
            RIDValue = TIMS.Get_RIDforOrgID(TIMS.Get_OrgIDforComIDNO(ComIDNOValue), sm.UserInfo.PlanID)
            'RIDValue = TIMS.Get_RIDforOrgID(TIMS.Get_OrgIDforComIDNO(ComIDNOValue), "1024")
            If RIDValue = Nothing Then
                Errmsg += " 此機構尚未加入此計劃" & vbCrLf
                Exit Function
            End If

            '準備儲存資料
            sql = "select * from Plan_PlanInfo where 1<>1"
            dt = DbAccess.GetDataTable(sql, da, Trans)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("PlanID") = sm.UserInfo.PlanID
            dr("ComIDNO") = ComIDNOValue 'colArray(Cst_ComIDNO).ToString
            dr("SeqNO") = int_SeqNO

            dr("RID") = RIDValue
            dr("PlanYear") = sm.UserInfo.Years
            dr("TPlanID") = sm.UserInfo.TPlanID

            dr("ClassName") = IIf(colArray(Cst_ClassName).ToString <> "", colArray(Cst_ClassName).ToString, Convert.DBNull)
            If colArray(Cst_GCID).ToString <> "" Then
                dr("GCID") = colArray(Cst_GCID).ToString
            Else
                dr("GCID") = Convert.DBNull
            End If
            dr("TMID") = IIf(colArray(Cst_TMID).ToString <> "", colArray(Cst_TMID).ToString, Convert.DBNull)

            If colArray(Cst_PlanEMail).ToString = "" Then
                dr("PlanEMail") = Convert.DBNull
            Else
                dr("PlanEMail") = colArray(Cst_PlanEMail).ToString
            End If

            dr("PlanCause") = IIf(colArray(Cst_PlanCause).ToString <> "", colArray(Cst_PlanCause).ToString, Convert.DBNull)
            dr("PurScience") = IIf(colArray(Cst_PurScience).ToString <> "", colArray(Cst_PurScience).ToString, Convert.DBNull)
            dr("PurTech") = IIf(colArray(Cst_PurTech).ToString <> "", colArray(Cst_PurTech).ToString, Convert.DBNull)
            dr("PurMoral") = IIf(colArray(Cst_PurMoral).ToString <> "", colArray(Cst_PurMoral).ToString, Convert.DBNull)

            dr("CapDegree") = IIf(colArray(Cst_CapDegree).ToString <> "", colArray(Cst_CapDegree).ToString, "00")
            dr("CapAge1") = IIf(colArray(Cst_CapAge1).ToString <> "", colArray(Cst_CapAge1).ToString, "15")
            dr("CapAge2") = IIf(colArray(Cst_CapAge2).ToString <> "", colArray(Cst_CapAge2).ToString, "65")
            dr("CapSex") = "0"
            dr("CapMilitary") = "00"

            If colArray(Cst_CapOther1).ToString = "" Then
                dr("CapOther1") = Convert.DBNull
            Else
                dr("CapOther1") = colArray(Cst_CapOther1).ToString
            End If
            If colArray(Cst_CapOther2).ToString = "" Then
                dr("CapOther2") = Convert.DBNull
            Else
                dr("CapOther2") = colArray(Cst_CapOther2).ToString
            End If
            If colArray(Cst_CapOther3).ToString = "" Then
                dr("CapOther3") = Convert.DBNull
            Else
                dr("CapOther3") = colArray(Cst_CapOther3).ToString
            End If

            dr("TMScience") = IIf(colArray(Cst_TMScience).ToString <> "", colArray(Cst_TMScience).ToString, Convert.DBNull)

            dr("GenSciHours") = Convert.DBNull
            dr("ProSciHours") = Convert.DBNull
            dr("ProTechHours") = Convert.DBNull
            dr("OtherHours") = Convert.DBNull
            dr("TotalHours") = Convert.DBNull
            dr("DefGovCost") = Convert.DBNull '經費來源-政府負擔$ --計算後重新統合
            dr("DefUnitCost") = Convert.DBNull '經費來源-單位負擔% (X)
            dr("DefStdCost") = Convert.DBNull '經費來源-學員負擔$ --計算後重新統合
            dr("ProcID") = "" 'ClassChar.SelectedValue
            'ProcID 2008 拿掉，因為完全沒有用到，寫了也是白寫  by amu 2008-01-14

            'If DefGovCost.Text = "" Then
            '    dr("DefGovCost") = Convert.DBNull
            'Else
            '    dr("DefGovCost") = DefGovCost.Text
            'End If
            'If DefUnitCost.Text = "" Then
            '    dr("DefUnitCost") = Convert.DBNull
            'Else
            '    dr("DefUnitCost") = DefUnitCost.Text
            'End If
            'If DefStdCost.Text = "" Then
            '    dr("DefStdCost") = Convert.DBNull
            'Else
            '    dr("DefStdCost") = DefStdCost.Text
            'End If

            If colArray(Cst_FirstSort).ToString <> "" Then
                dr("FirstSort") = colArray(Cst_FirstSort).ToString
            Else
                dr("FirstSort") = Convert.DBNull
            End If
            dr("PointYN") = IIf(colArray(Cst_PointYN).ToString = "Y", "Y", "N")
            dr("IsBusiness") = IIf(colArray(Cst_IsBusiness).ToString = "Y", "Y", "N")
            dr("EnterpriseName") = colArray(Cst_EnterpriseName).ToString
            dr("TNum") = IIf(colArray(Cst_TNum).ToString <> "", colArray(Cst_TNum).ToString, Convert.DBNull)
            dr("THours") = IIf(colArray(Cst_THours).ToString <> "", colArray(Cst_THours).ToString, Convert.DBNull)
            dr("STDate") = IIf(colArray(Cst_STDate).ToString <> "", colArray(Cst_STDate).ToString, Convert.DBNull)
            dr("FDDate") = IIf(colArray(Cst_FDDate).ToString <> "", colArray(Cst_FDDate).ToString, Convert.DBNull)

            dr("CyclType") = IIf(colArray(Cst_CyclType).ToString <> "", colArray(Cst_CyclType).ToString, Convert.DBNull)
            dr("ClassCount") = IIf(colArray(Cst_ClassCount).ToString <> "", colArray(Cst_ClassCount).ToString, Convert.DBNull)


            If colArray(Cst_TaddressZip).ToString <> "" And IsNumeric(colArray(Cst_TaddressZip).ToString) Then
                dr("TaddressZip") = colArray(Cst_TaddressZip).ToString
            Else
                dr("TaddressZip") = Convert.DBNull
            End If

            If colArray(Cst_TaddressZip2w).ToString <> "" And IsNumeric(colArray(Cst_TaddressZip2w).ToString) Then
                dr("TaddressZip2w") = colArray(Cst_TaddressZip2w).ToString
            Else
                dr("TaddressZip2w") = Convert.DBNull
            End If
            dr("TAddress") = IIf(colArray(Cst_TAddress).ToString <> "", colArray(Cst_TAddress).ToString, Convert.DBNull)

            dr("CredPoint") = IIf(colArray(Cst_CredPoint).ToString <> "", colArray(Cst_CredPoint).ToString, Convert.DBNull)
            dr("SciPlaceID") = IIf(colArray(Cst_SciPlaceID).ToString <> "", colArray(Cst_SciPlaceID).ToString, Convert.DBNull)
            dr("TechPlaceID") = IIf(colArray(Cst_TechPlaceID).ToString <> "", colArray(Cst_TechPlaceID).ToString, Convert.DBNull)

            dr("ConNum") = IIf(colArray(Cst_ConNum).ToString <> "", colArray(Cst_ConNum).ToString, Convert.DBNull)
            dr("ContactName") = IIf(colArray(Cst_ContactName).ToString <> "", colArray(Cst_ContactName).ToString, Convert.DBNull)
            dr("ContactPhone") = IIf(colArray(Cst_ContactPhone).ToString <> "", colArray(Cst_ContactPhone).ToString, Convert.DBNull)
            dr("ContactFax") = IIf(colArray(Cst_ContactFax).ToString <> "", colArray(Cst_ContactFax).ToString, Convert.DBNull)
            dr("ContactEmail") = IIf(colArray(Cst_ContactEmail).ToString <> "", colArray(Cst_ContactEmail).ToString, Convert.DBNull)
            dr("ClassCate") = IIf(colArray(Cst_ClassCate).ToString <> "", colArray(Cst_ClassCate).ToString, Convert.DBNull)

            If colArray(Cst_EnterSupplyStyle).ToString <> "" Then
                dr("EnterSupplyStyle") = colArray(Cst_EnterSupplyStyle).ToString
            Else
                dr("EnterSupplyStyle") = Convert.DBNull
            End If

            If colArray(Cst_Note).ToString = "" Then
                dr("Note") = Convert.DBNull
            Else
                dr("Note") = colArray(Cst_Note).ToString
            End If
            dr("AppliedResult") = Convert.DBNull '尚未審核通過
            dr("TransFlag") = "N" '未轉班
            dr("IsApprPaper") = "Y" '已通過
            dr("AppliedDate") = Now.Date
            dr("AppliedOrigin") = 1

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, Trans)

            'TC_01_014_add.aspx 開班計畫表資料維護作業
            sql = " delete Plan_VerReport "
            sql += " WHERE PlanID='" & sm.UserInfo.PlanID & "' "
            sql += " and ComIDNO='" & ComIDNOValue & "' "
            sql += " and SeqNo='" & int_SeqNO & "' "
            DbAccess.ExecuteNonQuery(sql, Trans)

            'Trans = DbAccess.BeginTrans(conn)
            sql = " SELECT * FROM Plan_VerReport WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, Trans)
            dr = dt.NewRow
            dt.Rows.Add(dr)

            dr("PlanID") = sm.UserInfo.PlanID
            dr("ComIDNO") = ComIDNOValue
            dr("SeqNo") = int_SeqNO
            dr("ClassID") = ""
            'dr("Times") = Me.Times.Text
            dr("TPeriod") = Convert.DBNull
            'dr("TPeriod") = IIf(colArray(Cst_TPeriod).ToString <> "", colArray(Cst_TPeriod).ToString, Convert.DBNull)
            dr("TrainDemain") = IIf(colArray(Cst_TrainDemain).ToString <> "", colArray(Cst_TrainDemain).ToString, Convert.DBNull)

            Dim TrainTarget_Text As String = ""
            TrainTarget_Text += "緣由：" & colArray(Cst_PlanCause).ToString + vbCrLf
            TrainTarget_Text += "學科：" & colArray(Cst_PurScience).ToString + vbCrLf
            TrainTarget_Text += "技能：" & colArray(Cst_PurTech).ToString + vbCrLf
            TrainTarget_Text += "品德：" & colArray(Cst_PurMoral).ToString

            dr("TrainTarget") = TrainTarget_Text

            dr("TeacherDesc") = IIf(colArray(Cst_TeacherDesc).ToString <> "", colArray(Cst_TeacherDesc).ToString, Convert.DBNull)
            dr("Domain") = Convert.DBNull           'Me.Domain.Text
            dr("CapAll") = IIf(colArray(Cst_CapAll).ToString <> "", colArray(Cst_CapAll).ToString, Convert.DBNull)
            dr("CostDesc") = IIf(colArray(Cst_Note).ToString <> "", colArray(Cst_Note).ToString, Convert.DBNull)
            dr("TrainMode") = "" 'IIf(colArray(Cst_TrainMode).ToString <> "", colArray(Cst_TrainMode).ToString, Convert.DBNull)
            dr("Content") = "" '已更新為條列式

            dr("RecDesc") = IIf(colArray(Cst_RecDesc).ToString <> "", colArray(Cst_RecDesc).ToString, Convert.DBNull)
            dr("LearnDesc") = IIf(colArray(Cst_LearnDesc).ToString <> "", colArray(Cst_LearnDesc).ToString, Convert.DBNull)
            dr("ActDesc") = IIf(colArray(Cst_ActDesc).ToString <> "", colArray(Cst_ActDesc).ToString, Convert.DBNull)
            dr("ResultDesc") = IIf(colArray(Cst_ResultDesc).ToString <> "", colArray(Cst_ResultDesc).ToString, Convert.DBNull)
            dr("OtherDesc") = IIf(colArray(Cst_OtherDesc).ToString <> "", colArray(Cst_OtherDesc).ToString, Convert.DBNull)
            dr("Recruit") = IIf(colArray(Cst_Recruit).ToString <> "", colArray(Cst_Recruit).ToString, Convert.DBNull)
            dr("Inspire") = IIf(colArray(Cst_Inspire).ToString <> "", colArray(Cst_Inspire).ToString, Convert.DBNull)
            dr("IsApprPaper") = "Y"

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, Trans)


            Dim dtTemp As DataTable
            Dim str_Rule As String
            str_Rule = "[訓練機構統一編號]='" & colArray(Cst_ComIDNO).ToString & "'"
            '97產學訓課程大綱
            sql = " SELECT * FROM Plan_TrainDesc WHERE 1<>1"
            dtTemp = DbAccess.GetDataTable(sql, da, Trans)

            dtTemp.Columns("PTDID").AutoIncrement = True
            dtTemp.Columns("PTDID").AutoIncrementSeed = -1
            dtTemp.Columns("PTDID").AutoIncrementStep = -1

            For Each drw As DataRow In dt2.Select(str_Rule, Nothing, DataViewRowState.CurrentRows)
                If drw(Cst_ClassName).ToString = colArray(Cst_ClassName).ToString Then
                    dr = dtTemp.NewRow
                    dtTemp.Rows.Add(dr)
                    dr("PlanID") = sm.UserInfo.PlanID
                    dr("ComIDNO") = ComIDNOValue
                    dr("SeqNO") = int_SeqNO
                    dr("STrainDate") = drw(Cst_STrainDate)
                    dr("ETrainDate") = drw(Cst_STrainDate)
                    dr("PName") = drw(Cst_PName)
                    dr("PHour") = drw(Cst_PHour)
                    dr("PCont") = drw(Cst_PCont)
                    dr("Classification1") = CInt(drw(Cst_Classification1))
                    dr("PTID") = TIMS.Get_PTID(drw(Cst_PTID), ComIDNOValue)
                    dr("TechID") = drw(Cst_TechID)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
            Next
            dt = dtTemp.Copy
            DbAccess.UpdateDataTable(dt, da, Trans)


            ''建立上課時間(Plan_OnClass)
            sql = " SELECT * FROM Plan_OnClass WHERE 1<>1"
            dtTemp = DbAccess.GetDataTable(sql, da, Trans)

            dtTemp.Columns("POCID").AutoIncrement = True
            dtTemp.Columns("POCID").AutoIncrementSeed = -1
            dtTemp.Columns("POCID").AutoIncrementStep = -1

            For Each drw As DataRow In dt3.Select(str_Rule, Nothing, DataViewRowState.CurrentRows)
                If drw(Cst_ClassName).ToString = colArray(Cst_ClassName).ToString Then
                    dr = dtTemp.NewRow
                    dtTemp.Rows.Add(dr)
                    dr("PlanID") = sm.UserInfo.PlanID
                    dr("ComIDNO") = ComIDNOValue
                    dr("SeqNO") = int_SeqNO
                    'dr("Weeks") = drw(Cst_Weeks) '星期x
                    Select Case drw(Cst_Weeks)
                        Case "1"
                            dr("Weeks") = "星期一"
                        Case "2"
                            dr("Weeks") = "星期二"
                        Case "3"
                            dr("Weeks") = "星期三"
                        Case "4"
                            dr("Weeks") = "星期四"
                        Case "5"
                            dr("Weeks") = "星期五"
                        Case "6"
                            dr("Weeks") = "星期六"
                        Case "7"
                            dr("Weeks") = "星期日"
                        Case Else
                            dr("Weeks") = ""
                    End Select
                    dr("Times") = drw(Cst_Times) '時間
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
            Next
            dt = dtTemp.Copy
            DbAccess.UpdateDataTable(dt, da, Trans)


            '計畫經費項目檔(Plan_CostItem)
            sql = " SELECT * FROM Plan_CostItem WHERE 1<>1"
            dtTemp = DbAccess.GetDataTable(sql, da, Trans)

            dtTemp.Columns("PCID").AutoIncrement = True
            dtTemp.Columns("PCID").AutoIncrementSeed = -1
            dtTemp.Columns("PCID").AutoIncrementStep = -1

            For Each drw As DataRow In dt4.Select(str_Rule, Nothing, DataViewRowState.CurrentRows)
                If drw(Cst_ClassName).ToString = colArray(Cst_ClassName).ToString Then
                    dr = dtTemp.NewRow
                    dtTemp.Rows.Add(dr)
                    dr("PlanID") = sm.UserInfo.PlanID
                    dr("ComIDNO") = ComIDNOValue
                    dr("SeqNO") = int_SeqNO
                    dr("CostMode") = 5
                    If drw(Cst_CostID).ToString.Length = 1 Then
                        dr("CostID") = "0" & drw(Cst_CostID) '項目
                    Else
                        dr("CostID") = drw(Cst_CostID) '項目
                    End If
                    dr("OPrice") = drw(Cst_OPrice) '單價
                    dr("Itemage") = drw(Cst_Itemage)  '計價數量
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
            Next
            dt = dtTemp.Copy
            DbAccess.UpdateDataTable(dt, da, Trans)

            DbAccess.CommitTrans(Trans)
            Save_Plan_PlanInfo = True
        Catch ex As Exception
            Errmsg += ex.Message
            DbAccess.RollbackTrans(Trans)
            Save_Plan_PlanInfo = False
        End Try

    End Function

    Function GetMaxSeqNum(ByVal Trans As SqlTransaction, ByVal ComidValue As String) As Integer
        Dim sql As String = ""
        Dim dr As DataRow
        '取得SeqNO
        sql = "SELECT PlanID,ComIDNO,SeqNO From Plan_PlanInfo where ComIDNO='" & ComidValue & "' and PlanID='" & sm.UserInfo.PlanID & "' order by SeqNO Desc"
        dr = DbAccess.GetOneRow(sql, Trans)
        If dr Is Nothing Then
            Return 1
        Else
            Return dr("SeqNO") + 1
        End If
    End Function



End Class
