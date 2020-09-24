VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Generic Try-Catch Block inside VB by Vanja Fuckar"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Test Boundary"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Copy Memory"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "(Works only with compiled EXE)"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim IsError As Long 'Reference Notification!

EnterTry IsError 'ENTER TRY - CATCH BLOCK

If IsError = 0 Then
    'TRY BLOCK
    CopyMemory ByVal 15&, ByVal &H3000&, 10000
Else
    'CATCH ERROR BLOCK
    MsgBox "Error inside TRY BLOCK!", vbExclamation, "Information"
End If

LeaveTry 'MUST LEAVE TRY-CATCH BLOCK
End Sub

Private Sub Command2_Click()
Dim IsError As Long 'Reference Notification!
Dim S() As String

EnterTry IsError 'ENTER TRY - CATCH BLOCK

If IsError = 0 Then
    'TRY BLOCK
     S(1) = "TEST ITEM"
    
Else
    'CATCH ERROR BLOCK
    ReDim S(1)
    S(1) = "TEST ITEM SUCCESS"

End If

LeaveTry 'MUST LEAVE TRY-CATCH BLOCK

MsgBox "Item:" & S(1) & vbCrLf & "Array Bound:" & UBound(S), vbExclamation, "Info"

End Sub

