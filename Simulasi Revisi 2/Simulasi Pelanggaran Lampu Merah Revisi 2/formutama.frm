VERSION 5.00
Begin VB.Form frmUtama 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TUGAS PROJECT PKTI"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   255
      Left            =   4560
      TabIndex        =   48
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Left            =   2160
      TabIndex        =   47
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   46
      Top             =   3960
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   45
      Top             =   1920
      Width           =   135
   End
   Begin VB.Timer tmrArah 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1800
   End
   Begin VB.PictureBox picArah 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   3240
      ScaleHeight     =   735
      ScaleWidth      =   375
      TabIndex        =   43
      Top             =   5040
      Width           =   375
      Begin VB.Label lblArahAnim 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   3
         Left            =   0
         TabIndex        =   44
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox picArah 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   480
      ScaleHeight     =   255
      ScaleWidth      =   1335
      TabIndex        =   41
      Top             =   2760
      Width           =   1335
      Begin VB.Label lblArahAnim 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "g"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picArah 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   3960
      ScaleHeight     =   615
      ScaleWidth      =   375
      TabIndex        =   39
      Top             =   720
      Width           =   375
      Begin VB.Label lblArahAnim 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   1
         Left            =   0
         TabIndex        =   40
         Top             =   -120
         Width           =   255
      End
   End
   Begin VB.PictureBox picArah 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   5760
      ScaleHeight     =   375
      ScaleWidth      =   1335
      TabIndex        =   37
      Top             =   3360
      Width           =   1335
      Begin VB.Label lblArahAnim 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   2
         Left            =   840
         TabIndex        =   38
         Top             =   -120
         Width           =   375
      End
   End
   Begin VB.Timer tmrAnim 
      Interval        =   100
      Left            =   240
      Top             =   6240
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   34
      Top             =   5160
      Width           =   2055
      Begin VB.TextBox txtHijau 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtKuning 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kuning =            Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hijau     =            Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   31
      Top             =   5160
      Width           =   2055
      Begin VB.TextBox txtHijau 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtKuning 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hijau     =            Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kuning =            Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   28
      Top             =   600
      Width           =   2055
      Begin VB.TextBox txtHijau 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtKuning 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kuning =             Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hijau     =             Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   25
      Top             =   720
      Width           =   1935
      Begin VB.TextBox txtHijau 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtKuning 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hijau     =            Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kuning =            Detik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1710
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   2520
      TabIndex        =   24
      Top             =   5880
      Width           =   2655
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000013&
         Caption         =   "&EXIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         MouseIcon       =   "formutama.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "formutama.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         BackColor       =   &H80000013&
         Caption         =   "&STOP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         MouseIcon       =   "formutama.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "formutama.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRun 
         BackColor       =   &H80000013&
         Caption         =   "&R U N"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MouseIcon       =   "formutama.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "formutama.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   13
         Left            =   0
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Timer tmrLampu 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   6360
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   2880
      Picture         =   "formutama.frx":15E4
      Top             =   720
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   600
      Picture         =   "formutama.frx":26196
      Top             =   3120
      Width           =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   5280
      X2              =   5280
      Y1              =   2520
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   2280
      X2              =   2280
      Y1              =   2520
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   3120
      X2              =   4560
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   3120
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   14
      Left            =   4200
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   12
      Left            =   3240
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   11
      Left            =   3960
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   10
      Left            =   3480
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   15
      Left            =   4680
      Top             =   3600
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   14
      Left            =   4680
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   13
      Left            =   4680
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   12
      Left            =   4680
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   11
      Left            =   2280
      Top             =   3600
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   10
      Left            =   2280
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   9
      Left            =   4200
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   3960
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   3240
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   3480
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   9
      Left            =   2280
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   8
      Left            =   2280
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   2520
      MouseIcon       =   "formutama.frx":27298
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   4560
      MouseIcon       =   "formutama.frx":275A2
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   5520
      MouseIcon       =   "formutama.frx":278AC
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   2640
      MouseIcon       =   "formutama.frx":27BB6
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   2040
      MouseIcon       =   "formutama.frx":27EC0
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   5040
      MouseIcon       =   "formutama.frx":281CA
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   4560
      MouseIcon       =   "formutama.frx":284D4
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   2280
      MouseIcon       =   "formutama.frx":287DE
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblLampuMerah 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   5760
      MouseIcon       =   "formutama.frx":28AE8
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblLampuMerah 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   5520
      MouseIcon       =   "formutama.frx":28DF2
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblLampuMerah 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   4680
      MouseIcon       =   "formutama.frx":290FC
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblLampuMerah 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   1920
      MouseIcon       =   "formutama.frx":29406
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   3720
      Top             =   5160
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   3720
      Top             =   4560
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   3720
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   3720
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   3720
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   3720
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   7
      Left            =   6840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   6120
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   5400
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   4680
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   120
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   2280
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   1560
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblJudul 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "SIMULASI PELANGGARAN LAMPU MERAH"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   5250
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5295
      Left            =   3000
      Top             =   600
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   120
      Top             =   2520
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   600
      Picture         =   "formutama.frx":29710
      Top             =   2640
      Width           =   720
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function Inp Lib "inpout32.dll" _
Alias "Inp32" (ByVal PortAddress As Integer) As Integer

Private Declare Sub Out Lib "inpout32.dll" _
Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

Dim pantul As Integer
Dim idxLampuHijau As Integer

Private Sub LampuMati()
    Dim ctl As Control
    
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is Shape Then
            If ctl.Name = "shpLampuMerah" Then ctl.BackColor = RGB(70, 0, 0)
            If ctl.Name = "shpLampuKuning" Then ctl.BackColor = RGB(70, 70, 0)
            If ctl.Name = "shpLampuHijau" Then ctl.BackColor = RGB(0, 70, 0)
        End If
    Next
End Sub

Private Sub LampuMerahNyala(Index As Integer)
    Select Case Index
    Case 0
        
    Case 1
        
    Case 2
        
    Case 3
        
    End Select
    
    shpLampuMerah(Index).BackColor = vbRed
End Sub

Private Sub LampuMerahMati(Index As Integer)
    Select Case Index
    Case 0
        Out &H378, Val(Inp(&H378)) - 1
    Case 1
        Out &H378, Val(Inp(&H378)) - 2
    Case 2
        Out &H378, Val(Inp(&H378)) - 4
    Case 3
        Out &H378, Val(Inp(&H378)) - 8
    End Select
    
    shpLampuMerah(Index).BackColor = RGB(50, 0, 0)
End Sub

Private Sub LampuKuningNyala(Index As Integer)
    Select Case Index
    Case 0
        
    Case 1

    Case 2

    Case 3

    End Select
    
    shpLampuKuning(Index).BackColor = vbYellow
End Sub

Private Sub LampuKuningMati(Index As Integer)
    Select Case Index
    Case 0
        Out &H378, Val(Inp(&H378)) - 16
    Case 1
        Out &H378, Val(Inp(&H378)) - 32
    Case 2
        Out &H378, Val(Inp(&H378)) - 64
    Case 3
        Out &H378, Val(Inp(&H378)) - 128
    End Select
    
    shpLampuKuning(Index).BackColor = RGB(50, 50, 0)
End Sub

Private Sub LampuHijauNyala(Index As Integer)
    ResetArahAnim
    Select Case Index
    Case 0
        
        idxLampuHijau = 0
    Case 1
        
        idxLampuHijau = 1
    Case 2
        
        idxLampuHijau = 2
    Case 3
        
        idxLampuHijau = 3
    End Select
    shpLampuHijau(Index).BackColor = vbGreen
    tmrArah.Enabled = True
End Sub

Private Sub LampuHijauMati(Index As Integer)
    tmrArah.Enabled = False
    ResetArahAnim
    Select Case Index
    Case 0
        Out &H37A, 11
    Case 1
        Out &H37A, 11
    Case 2
        Out &H37A, 11
    Case 3
        Out &H37A, 11
    End Select
    shpLampuHijau(Index).BackColor = RGB(0, 50, 0)
End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdRun_Click()
    Dim intNum As Integer
    
    LampuMati
    tmrLampu.Interval = 1
    tmrLampu.Enabled = True
End Sub

Private Sub cmdStop_Click()
    tmrArah.Enabled = False
    LampuMati
    tmrLampu.Enabled = False
End Sub

Private Sub ResetArahAnim()
    With lblArahAnim(0)
        .Move 0 - .Width, (picArah(0).ScaleHeight - .Height) / 2
    End With
    With lblArahAnim(1)
        .Move (picArah(1).ScaleWidth - .Width) / 8, 0 - .Height
    End With
    With lblArahAnim(2)
        .Move picArah(2).ScaleWidth + .Width, (picArah(2).ScaleHeight - .Height) / 2
    End With
    With lblArahAnim(3)
        .Move (picArah(3).ScaleWidth - .Width) / 8, picArah(3).ScaleHeight + .Height
    End With
End Sub

Private Sub Form_Load()
    ResetArahAnim
    LampuMati
    blnHijau = True
    blnKuning = False
    blnMerah = False
    pantul = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LampuMati
End Sub

Private Sub lblLampuHijau_Click(Index As Integer)
    LampuMati
    LampuHijauNyala Index
End Sub

Private Sub lblLampuHijau_DblClick(Index As Integer)
    LampuHijauMati Index
End Sub

Private Sub lblLampuKuning_Click(Index As Integer)
    LampuMati
    LampuKuningNyala Index
End Sub

Private Sub lblLampuKuning_DblClick(Index As Integer)
    LampuKuningMati Index
End Sub

Private Sub lblLampuMerah_Click(Index As Integer)
    LampuMati
    LampuMerahNyala Index
End Sub

Private Sub lblLampuMerah_DblClick(Index As Integer)
    LampuMerahMati Index
End Sub

Private Sub tmrAnim_Timer()
    With lblJudul
        .Left = .Left + pantul
        If .Left < 0 Then pantul = 100
        If .Left > Me.ScaleWidth - .Width Then pantul = -100
    End With
    
End Sub

Private Sub tmrArah_Timer()
    With lblArahAnim(idxLampuHijau)
        Select Case idxLampuHijau
        Case 0
            .Left = .Left + 20
            If .Left > picArah(idxLampuHijau).ScaleWidth Then .Left = 0 - .Width
        Case 1
            .Top = .Top + 20
            If .Top > picArah(idxLampuHijau).ScaleHeight Then .Top = 0 - .Height
        Case 2
            .Left = .Left - 20
            If .Left < 0 - .Width Then .Left = picArah(idxLampuHijau).ScaleWidth
        Case 3
            .Top = .Top - 20
            If .Top < 0 - .Height Then .Top = picArah(idxLampuHijau).ScaleHeight
        End Select
    End With
End Sub

Private Sub tmrLampu_Timer()
    Static Index As Integer
    Static intLampu As Integer
    Dim intNum As Integer
    
    Select Case intLampu
    Case 0 'Hijau
        LampuMati
        tmrLampu.Interval = Val(txtHijau(Index).Text) * 1000
        LampuHijauNyala Index
        For intNum = 0 To 3
            If intNum <> Index Then LampuMerahNyala intNum
        Next
        intLampu = 1
    Case 1 'Kuning
        LampuMati
        tmrLampu.Interval = Val(txtKuning(Index).Text) * 1000
        LampuKuningNyala Index
        For intNum = 0 To 3
            If intNum <> Index Then LampuMerahNyala intNum
        Next
        intLampu = 0
        Index = Index + 1
        If Index = 4 Then Index = 0
    End Select
End Sub

Private Sub txtHijau_Change(Index As Integer)
    With txtHijau(Index)
        If IsNumeric(.Text) = False Then SendKeys vbBack: Exit Sub
    End With
End Sub

Private Sub txtKuning_Change(Index As Integer)
    With txtKuning(Index)
        If IsNumeric(.Text) = False Then SendKeys vbBack: Exit Sub
    End With
End Sub

