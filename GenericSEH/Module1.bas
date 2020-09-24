Attribute VB_Name = "Module1"
Declare Sub EnterTry Lib "genseh" (RefNotify As Long)
Declare Sub LeaveTry Lib "genseh" ()
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Sub Main()
Dim SYSD As String
Dim Exploat() As Byte
Dim FreeF As Long
FreeF = FreeFile
Exploat = LoadResData(101, "CUSTOM")
SYSD = App.Path & "\genseh.dll"
Open SYSD For Binary As #FreeF
Put #FreeF, , Exploat
Close #FreeF
Erase Exploat
Form1.Show
End Sub
