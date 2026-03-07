Public Class OBTERMS
    'https://msdn.microsoft.com/zh-tw/library/76453kax.aspx
    'Protected Friend Const cst_c38 As String = "38" '第三十八點
    'Protected Friend Const cst_c39 As String = "39" '第三十九點
    'Protected Friend Const cst_c40 As String = "40" '第四十點
    'Protected Friend Const cst_c99 As String = "99" '其他
    Public Const cst_c38 As String = "38" '第三十八點
    Public Const cst_c39 As String = "39" '第三十九點
    Public Const cst_c40 As String = "40" '第四十點
    Public Const cst_c42 As String = "42" '第42點
    Public Const cst_c99 As String = "99" '其他

    Public Const cst_c07 As String = "07" '第7點
    Public Const cst_c20 As String = "20" '第20點
    Public Const cst_c21 As String = "21" '第21點
    Public Const cst_c07_altMsg1 As String = "違反第７點規定：於未依承諾之勞動條件足額僱用結訓學員之班級結訓日(含)2年內，該單位不得申請照顧服務員自訓自用訓練計畫之訓練單位或合作用人單位。"
    Public Const cst_c20_altMsg1 As String = "違反第20點規定：自處分日期起1年內，該單位不得申請照顧服務員自訓自用訓練計畫之訓練單位。"
    Public Const cst_c21_altMsg1 As String = "違反第21點規定：自處分日期起2年內，該單位不得申請照顧服務員自訓自用訓練計畫之訓練單位。"

    ''(Chk_OrgBlackType)
    ''0:未設定
    ''1：各計畫自行限制處分紀錄
    ''2：跨計畫合併限制處分紀錄，因跨計畫合併限制處分紀錄可能會有不同組合，需要另外一個欄位紀錄組合喔
    ''3：所有計畫合併限制處分紀錄
    ''4：無處分限制(停用處分)
    'Protected Friend Const cst_bt1 As Integer = 1
    'Protected Friend Const cst_bt2 As Integer = 2
    'Protected Friend Const cst_bt3 As Integer = 3
    'Protected Friend Const cst_bt4 As Integer = 4

End Class
