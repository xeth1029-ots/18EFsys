'------------------------------------------------------------------------------
' <自動產生的>
'     這段程式碼是由工具產生的。
'
'     變更這個檔案可能會導致不正確的行為，而且如果已重新產生
'     程式碼，則會遺失變更。
' </自動產生的>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Partial Public Class TC_01_017_add
    
    '''<summary>
    '''form1 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents form1 As Global.System.Web.UI.HtmlControls.HtmlForm
    
    '''<summary>
    '''TitleLab1 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents TitleLab1 As Global.System.Web.UI.WebControls.Label
    
    '''<summary>
    '''TitleLab2 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents TitleLab2 As Global.System.Web.UI.WebControls.Label
    
    '''<summary>
    '''tb_orgname 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents tb_orgname As Global.System.Web.UI.WebControls.TextBox
    
    '''<summary>
    '''hid_orgid 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents hid_orgid As Global.System.Web.UI.HtmlControls.HtmlInputHidden
    
    '''<summary>
    '''tb_comidno 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents tb_comidno As Global.System.Web.UI.WebControls.TextBox
    
    '''<summary>
    '''dl_typeid1 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents dl_typeid1 As Global.System.Web.UI.WebControls.DropDownList
    
    '''<summary>
    '''dl_typeid2 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents dl_typeid2 As Global.System.Web.UI.WebControls.DropDownList
    
    '''<summary>
    '''city_code 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents city_code As Global.System.Web.UI.HtmlControls.HtmlInputText
    
    '''<summary>
    '''ZipCODEB3 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents ZipCODEB3 As Global.System.Web.UI.HtmlControls.HtmlInputText
    
    '''<summary>
    '''hidZipCODE6W 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents hidZipCODE6W As Global.System.Web.UI.HtmlControls.HtmlInputHidden
    
    '''<summary>
    '''Litcity_code 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents Litcity_code As Global.System.Web.UI.WebControls.Literal
    
    '''<summary>
    '''TBCity 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents TBCity As Global.System.Web.UI.WebControls.TextBox
    
    '''<summary>
    '''city_zip 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents city_zip As Global.System.Web.UI.HtmlControls.HtmlInputButton
    
    '''<summary>
    '''TBaddress 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents TBaddress As Global.System.Web.UI.WebControls.TextBox
    
    '''<summary>
    '''rfvcity 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents rfvcity As Global.System.Web.UI.WebControls.RequiredFieldValidator
    
    '''<summary>
    '''rfvaddress 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents rfvaddress As Global.System.Web.UI.WebControls.RequiredFieldValidator
    
    '''<summary>
    '''bt_save 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents bt_save As Global.System.Web.UI.WebControls.Button
    
    '''<summary>
    '''bt_back 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents bt_back As Global.System.Web.UI.WebControls.Button
    
    '''<summary>
    '''msg 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents msg As Global.System.Web.UI.WebControls.Label
End Class
