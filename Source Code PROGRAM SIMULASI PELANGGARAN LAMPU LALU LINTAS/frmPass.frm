VERSION 5.00
Begin VB.Form frmPass 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "                  Isikan Password !!!"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3240
   ControlBox      =   0   'False
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "&BATAL"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MouseIcon       =   "frmPass.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "="
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Static Kesempatan As Integer
    
    If KeyAscii <> 13 Then Exit Sub
    
    With txtPass
        If .Text = "amir" Then
            Unload Me
            frmUtama.Show
        Else
            MsgBox "PASSWORD YANG ANDA ISIKAN SALAH...", vbCritical, "IZIN AKSES DITOLAK"
            .Text = ""
            .SetFocus
            Kesempatan = Kesempatan + 1
            If Kesempatan = 3 Then
                MsgBox "MAAF!! ANDA TIDAK BERHAK MENGGUNAKAN PROGRAM INI...", vbCritical, "PROGRAM DITUTUP"
                Kesempatan = 0: End
            End If
        End If
    End With
End Sub
