VERSION 5.00
Begin VB.Form frmUtama 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TUGAS PROJECT PKTI"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13365
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
   ScaleHeight     =   8835
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   255
      Left            =   8040
      TabIndex        =   47
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Left            =   5040
      TabIndex        =   46
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Index           =   0
      Left            =   7800
      TabIndex        =   45
      Top             =   2280
      Width           =   255
   End
   Begin VB.Timer tmrArah 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   2880
      Top             =   1560
   End
   Begin VB.PictureBox picArah 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   7200
      ScaleHeight     =   735
      ScaleWidth      =   375
      TabIndex        =   43
      Top             =   6000
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
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   1335
      TabIndex        =   41
      Top             =   3720
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
         Left            =   840
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
      Left            =   7200
      ScaleHeight     =   615
      ScaleWidth      =   375
      TabIndex        =   39
      Top             =   720
      Width           =   375
      Begin VB.Label lblArahAnim 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "o"
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
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picArah 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   9840
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   37
      Top             =   4440
      Width           =   975
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
         Left            =   360
         TabIndex        =   38
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Timer tmrAnim 
      Interval        =   100
      Left            =   3120
      Top             =   7320
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
      Left            =   10800
      TabIndex        =   34
      Top             =   6000
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
      Left            =   2160
      TabIndex        =   31
      Top             =   6120
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
      Left            =   10920
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
      Left            =   2280
      TabIndex        =   25
      Top             =   480
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
      Left            =   5640
      TabIndex        =   24
      Top             =   7920
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
         MouseIcon       =   "revisi 2_2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "revisi 2_2.frx":030A
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
         MouseIcon       =   "revisi 2_2.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "revisi 2_2.frx":0A56
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
         MouseIcon       =   "revisi 2_2.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "revisi 2_2.frx":11A2
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
      Left            =   11640
      Top             =   7200
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   22
      Left            =   720
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   21
      Left            =   1680
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   20
      Left            =   4320
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   19
      Left            =   3480
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   15
      Left            =   2520
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   27
      Left            =   6840
      Top             =   7320
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   26
      Left            =   6840
      Top             =   6600
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   22
      Left            =   6840
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   25
      Left            =   7320
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   24
      Left            =   7560
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   23
      Left            =   7320
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   21
      Left            =   7560
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   20
      Left            =   6360
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   19
      Left            =   6120
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   18
      Left            =   7080
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   17
      Left            =   6600
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   16
      Left            =   6840
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   18
      Left            =   10920
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   17
      Left            =   11880
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   16
      Left            =   12720
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   15
      Left            =   6840
      Top             =   1680
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   6000
      Picture         =   "revisi 2_2.frx":15E4
      Top             =   360
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   840
      Picture         =   "revisi 2_2.frx":26196
      Top             =   4080
      Width           =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   8760
      X2              =   8760
      Y1              =   3480
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   5160
      X2              =   5160
      Y1              =   3480
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   6000
      X2              =   7800
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   14
      Left            =   8160
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   13
      Left            =   8160
      Top             =   3840
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   12
      Left            =   8160
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   11
      Left            =   5160
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   10
      Left            =   5160
      Top             =   3600
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   9
      Left            =   6840
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   6600
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   7080
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   6120
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   9
      Left            =   5160
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   8
      Left            =   5160
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   2520
      MouseIcon       =   "revisi 2_2.frx":27298
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   8040
      MouseIcon       =   "revisi 2_2.frx":275A2
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   5520
      MouseIcon       =   "revisi 2_2.frx":278AC
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblLampuHijau 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   5160
      MouseIcon       =   "revisi 2_2.frx":27BB6
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   2040
      MouseIcon       =   "revisi 2_2.frx":27EC0
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   8400
      MouseIcon       =   "revisi 2_2.frx":281CA
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   15
      Index           =   1
      Left            =   9000
      MouseIcon       =   "revisi 2_2.frx":284D4
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblLampuKuning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   4800
      MouseIcon       =   "revisi 2_2.frx":287DE
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblLampuMerah 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   5760
      MouseIcon       =   "revisi 2_2.frx":28AE8
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
      MouseIcon       =   "revisi 2_2.frx":28DF2
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblLampuMerah 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   8880
      MouseIcon       =   "revisi 2_2.frx":290FC
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lblLampuMerah 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   4560
      MouseIcon       =   "revisi 2_2.frx":29406
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape shpLampuHijau 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpLampuKuning 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape shpLampuMerah 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   6840
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   6840
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   6360
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   7
      Left            =   10080
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   9120
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   8160
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   8160
      Top             =   3600
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   5160
      Top             =   4080
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
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   5250
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   8895
      Left            =   6000
      Top             =   0
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1695
      Index           =   0
      Left            =   120
      Top             =   3360
      Width           =   13815
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function Inp Lib "input.dll" _
Alias "Inp32" (ByVal PortAddress As Integer) As Integer

Private Declare Sub Out Lib "input32.dll" _
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

