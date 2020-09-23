VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "What's My Processor?"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   435
      Left            =   4500
      TabIndex        =   2
      Top             =   1560
      Width           =   1275
   End
   Begin VB.CommandButton CmdGetProc 
      Caption         =   "Whats My Processor?"
      Height          =   435
      Left            =   1380
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label LblProc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   1140
      Width           =   5475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Function GetCPUDescription() As String

    Dim HKey As Long
    Dim C As Long
    Dim R As Long
    Dim S As String
    Dim T As Long
    
    R = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", 0, KEY_READ, HKey)
    C = 255
    S = String(C, Chr(0))
    R = RegQueryValueEx(HKey, "ProcessorNameString", 0, T, S, C)
    
    
    GetCPUDescription = Trim(Left(S, C - 1))

End Function


Private Sub cmdclose_Click()

    End
    
End Sub

Private Sub CmdGetProc_Click()

    LblProc.Caption = GetCPUDescription
    
End Sub
