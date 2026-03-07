Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' 一般的組件資訊是由下列這組屬性所控制。
' 變更這些屬性的值即可修改組件的相關資訊。

' 檢閱組件屬性的值

<Assembly: AssemblyTitle("")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("")> 
<Assembly: AssemblyCopyright("")> 
<Assembly: AssemblyTrademark("")> 
<Assembly: CLSCompliant(True)> 

'下列 GUID 為專案公開 (Expose) 至 COM 時所要使用的 typelib ID
<Assembly: Guid("26D0842A-B606-4DB1-9A71-A4685705DAEE")> 

' 組件的版本資訊由下列四個值所組成:
'
'      主要版本
'      次要版本
'      組建編號
'      修訂
'
' 您可以自行指定所有的值，也可以依照以下的方式，使用 '*' 將修訂和組建編號
' 指定為預設值:

<Assembly: AssemblyVersion("1.0.*")> 
