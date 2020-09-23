Attribute VB_Name = "PropBagCompile"
'                    SaLiLoG 1.0
'                  By SaLar Zeynali
'        Salixem@Gmail.Com - S4LiX3M@Yahoo.Com
'  _________      .____    .______  ___        _____
' /   _____/____  |    |   |__\   \/  /____   /     \
' \_____  \\__  \ |    |   |  |\     // __ \ /  \ /  \
' /        \/ __ \|    |___|  |/     \  ___//    Y    \
'/_______  (____  /_______ \__/___/\  \___  >____|__  /
'        \/     \/        \/        \_/   \/        \/

Public Function LoadCompiledData(Optional FileofData As String = "appexe") As PropertyBag
On Error GoTo Quit:
If FileofData = "appexe" Then: FileofData = App.Path & "\" & App.EXEName & ".exe"
Set LoadCompiledData = New PropertyBag
Dim DataBeginPos As Long
Dim varTemp As Variant
Dim byteArr() As Byte
Open FileofData For Binary As #1
Get #1, LOF(1) - 3, DataBeginPos
Seek #1, DataBeginPos
Get #1, , varTemp
byteArr = varTemp
LoadCompiledData.Contents = byteArr
Close #1
LoadCompiledData.WriteProperty "error", "false"
Exit Function
Quit:
Close #1
LoadCompiledData.WriteProperty "error", "true"
End Function
Public Sub CompileData(StubFile As String, PropBag As PropertyBag)
Dim DataBeginPos As Long
Dim varTemp As Variant
Open StubFile For Binary As #1
DataBeginPos = LOF(1)
varTemp = PropBag.Contents
Seek #1, LOF(1)
Put #1, , varTemp
Put #1, , DataBeginPos
Close #1
End Sub
