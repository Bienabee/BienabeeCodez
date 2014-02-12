Sub ConnectToQC(ByVal qcserver_txt, ByVal qcport_txt, ByVal qcusername_txt, ByVal qcpwd_txt, ByVal qcdomain_txt, ByValqcproj_txt)
 
Dim tdc1 As TDAPIOLELib.TDConnection
tdc1 =
New TDAPIOLELib.TDConnection
tdc1.InitConnectionEx(
"http://" & qcserver_txt & ":" & qcport_txt & "/qcbin")
tdc1.Login(qcusername_txt, qcpwd_txt)
tdc1.Connect(qcdomain_txt, qcproj_txt)
DimBugF
 
DimBugBg
BugF = tdc1.BugFactory()
BugBg = BugF.AddItem(System.
DBNull.Value)
BugBg.Summary =
"My Test Bug From VB App"
BugBg.Field(
"BG_DETECTION_DATE") = Today
BugBg.Field(
"BG_SEVERITY") = "2-Medium"
BugBg.Post()
　
　
If tdc1.Connected Then
MsgBox(
"You are connected to QC and created a Bug")
 
EndIf
