VERSION 5.00
Begin VB.Form frmBioRhythm 
   BackColor       =   &H00CB12A2&
   Caption         =   "  ~ KymaSoft Free BioRhythms ~ "
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7350
   FillStyle       =   0  'Solid
   Icon            =   "freebio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "freebio.frx":030A
   ScaleHeight     =   5040
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbReportYear 
      Height          =   315
      Left            =   6480
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox cmbReportDay 
      Height          =   315
      Left            =   5900
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.ComboBox cmbReportMonth 
      Height          =   315
      Left            =   4720
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbBirthYear 
      Height          =   315
      Left            =   3900
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox cmbBirthDay 
      Height          =   315
      Left            =   3320
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.ComboBox cmbBirthMonth 
      Height          =   315
      Left            =   2140
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox bioPic 
      BackColor       =   &H00C0E0FF&
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   128
      ScaleHeight     =   200
      ScaleMode       =   0  'User
      ScaleWidth      =   29
      TabIndex        =   8
      Top             =   1440
      Width           =   7095
      Begin VB.Line Line20 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   189.981
         Y2              =   189.981
      End
      Begin VB.Line Line19 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   180.039
         Y2              =   180.039
      End
      Begin VB.Line Line18 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29
         Y1              =   170.019
         Y2              =   170.019
      End
      Begin VB.Line Line16 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Line Line15 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   149.981
         Y2              =   149.981
      End
      Begin VB.Line Line14 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   139.961
         Y2              =   139.961
      End
      Begin VB.Line Line13 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   130.019
         Y2              =   130.019
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line11 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   109.981
         Y2              =   109.981
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   90.019
         Y2              =   90.019
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   80
         Y2              =   80
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   69.981
         Y2              =   69.981
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   59.961
         Y2              =   59.961
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   50.019
         Y2              =   50.019
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   40
         Y2              =   40
      End
      Begin VB.Line Line4 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.186
         Y1              =   29.981
         Y2              =   29.981
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29
         Y1              =   20.039
         Y2              =   20.039
      End
      Begin VB.Line Line1 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29
         Y1              =   10.019
         Y2              =   10.019
      End
      Begin VB.Line lnVert 
         X1              =   14.345
         X2              =   14.345
         Y1              =   0
         Y2              =   212.717
      End
      Begin VB.Line lnHor 
         X1              =   0
         X2              =   29.186
         Y1              =   101.734
         Y2              =   101.734
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   21
         X1              =   28.691
         X2              =   28.691
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   20
         X1              =   26.65
         X2              =   26.65
         Y1              =   0
         Y2              =   203.468
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   19
         X1              =   24.635
         X2              =   24.635
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   18
         X1              =   22.623
         X2              =   22.623
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   17
         X1              =   20.611
         X2              =   20.611
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   16
         X1              =   18.6
         X2              =   18.6
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   12
         X1              =   12.565
         X2              =   12.565
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   11
         X1              =   10.553
         X2              =   10.553
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   10
         X1              =   8.541
         X2              =   8.541
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   9
         X1              =   6.53
         X2              =   6.53
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   8
         X1              =   4.518
         X2              =   4.518
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   7
         X1              =   2.506
         X2              =   2.506
         Y1              =   0
         Y2              =   194.22
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   6
         X1              =   0.495
         X2              =   0.495
         Y1              =   0
         Y2              =   203.468
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   0
         X1              =   14.502
         X2              =   14.502
         Y1              =   0
         Y2              =   203.468
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   1
         X1              =   16.588
         X2              =   16.588
         Y1              =   0
         Y2              =   194.22
      End
   End
   Begin VB.TextBox userName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblCalcBtn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image imgCalcBtn 
      Height          =   480
      Left            =   2880
      Picture         =   "freebio.frx":6A0B
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1500
   End
   Begin VB.Line Line17 
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   7200
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line pipHor 
      Index           =   1
      X1              =   3840
      X2              =   3960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3075
      TabIndex        =   44
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   28
      Left            =   6960
      TabIndex        =   43
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   27
      Left            =   6735
      TabIndex        =   42
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   26
      Left            =   6495
      TabIndex        =   41
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   25
      Left            =   6240
      TabIndex        =   40
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   24
      Left            =   6000
      TabIndex        =   39
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   23
      Left            =   5760
      TabIndex        =   38
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   22
      Left            =   5505
      TabIndex        =   37
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   21
      Left            =   5265
      TabIndex        =   36
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   20
      Left            =   5025
      TabIndex        =   35
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   19
      Left            =   4770
      TabIndex        =   34
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   18
      Left            =   4530
      TabIndex        =   33
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   17
      Left            =   4290
      TabIndex        =   32
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   16
      Left            =   4035
      TabIndex        =   31
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   15
      Left            =   3795
      TabIndex        =   30
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   14
      Left            =   3555
      TabIndex        =   29
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   13
      Left            =   3300
      TabIndex        =   28
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   12
      Left            =   3060
      TabIndex        =   27
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   11
      Left            =   2820
      TabIndex        =   26
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   10
      Left            =   2565
      TabIndex        =   25
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   9
      Left            =   2325
      TabIndex        =   24
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   2085
      TabIndex        =   23
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   7
      Left            =   1830
      TabIndex        =   22
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   6
      Left            =   1590
      TabIndex        =   21
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   5
      Left            =   1350
      TabIndex        =   20
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   1095
      TabIndex        =   19
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   3
      Left            =   855
      TabIndex        =   18
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   615
      TabIndex        =   17
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label rptDateLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Report Date"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label nameLabel 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Intuitional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Intellectual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4160
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Emotional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2200
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Physical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "frmBioRhythm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Based on the core algorithm from a previous biorhythm
'program posted on Planet Source Code. I have corrected
'an error in my previous post that reversed the label for
'the emotional and intellectual cycles and added error
'handling to trap out of range date entries. None of the
'subroutines make calls that have the potential for
'changing the system date.

'The report date may be "quick-shifted" forward or back
'in time by clicking on the appropriate date beneath the
'chart.

'Biorhythm study and use is considered a "pseudoscience"
'in the United State however it is widely accepted and
'utilized throughout Europe and much of the rest of the
'world as a valid tool for guaging personal capabilities
'and performance based on a group of cycles that begin
'on the day of birth and extend throughout life.

'The code is fully functional as is but could be improved
'by the addition of a database for names and birthdates.
'Printed reports and generation of compatibility charts
'and/or "esoteric" charts remain to be implemented.

'Please email any suggestion for improving/enhancing the
'code to Michael Hebert at runes2@concentric.net or post
'them on Planet Source Code.

Option Explicit

Dim user As String
Dim birthMonth As String
Dim birthDay As String
Dim birthYear As String
Dim birthDate As String
Dim reportMonth As String
Dim reportDay As String
Dim reportYear As String
Dim reportDate As String
Dim timeWarp As Integer
Dim bioPeriod As Integer
Dim bioNum As Integer
Dim i As Integer
Dim n As Long

Private Sub Form_Load()

    Me.Show
    frmBioRhythm.Caption = " ~  KymaSoft Free BioRhythms  ~  Today is " & Date & "  ~"
    
    'Set ScaleWidth and ScaleHeight of picture box
    bioPic.ScaleWidth = 29 '14 days each side of report date
    bioPic.ScaleHeight = 200
        
    'Makes Horizontal line in the middle of the picture
    lnHor.Y1 = bioPic.ScaleHeight / 2
    lnHor.Y2 = lnHor.Y1
    lnHor.X1 = 0
    lnHor.X2 = bioPic.ScaleWidth
    
    'Makes Vertical line in the middle of the picture
    lnVert.Y1 = 0
    lnVert.Y2 = bioPic.ScaleHeight
    lnVert.X1 = bioPic.ScaleWidth / 2
    lnVert.X2 = bioPic.ScaleWidth / 2
    
    'Set default Report Date to current System Date
    reportDate = Date
    cmbReportMonth.Text = DatePart("m", reportDate)
            If cmbReportMonth.Text = "1" Then cmbReportMonth.Text = "January"
            If cmbReportMonth.Text = "2" Then cmbReportMonth.Text = "February"
            If cmbReportMonth.Text = "3" Then cmbReportMonth.Text = "March"
            If cmbReportMonth.Text = "4" Then cmbReportMonth.Text = "April"
            If cmbReportMonth.Text = "5" Then cmbReportMonth.Text = "May"
            If cmbReportMonth.Text = "6" Then cmbReportMonth.Text = "June"
            If cmbReportMonth.Text = "7" Then cmbReportMonth.Text = "July"
            If cmbReportMonth.Text = "8" Then cmbReportMonth.Text = "August"
            If cmbReportMonth.Text = "9" Then cmbReportMonth.Text = "September"
            If cmbReportMonth.Text = "10" Then cmbReportMonth.Text = "October"
            If cmbReportMonth.Text = "11" Then cmbReportMonth.Text = "November"
            If cmbReportMonth.Text = "12" Then cmbReportMonth.Text = "December"
    cmbReportDay.Text = DatePart("d", reportDate)
    cmbReportYear.Text = DatePart("yyyy", reportDate)
    
    fillBoxes
    userName_Change
    fillLabels
        
End Sub

Private Sub Form_Paint()

    'Repaints the Form and BioRhythm Chart if form has been
    'minimized or overlaid by another application
    
    If cmbBirthMonth.Text <> "" And cmbBirthDay.Text <> "" And cmbBirthYear.Text <> "" Then
        calc
    End If
    
End Sub

Private Sub cmdExit()

    Unload Me
    
End Sub

Private Sub fillBoxes()

    'Fills Birth Month ComboBox with Months
    cmbBirthMonth.AddItem "January"
    cmbBirthMonth.AddItem "February"
    cmbBirthMonth.AddItem "March"
    cmbBirthMonth.AddItem "April"
    cmbBirthMonth.AddItem "May"
    cmbBirthMonth.AddItem "June"
    cmbBirthMonth.AddItem "July"
    cmbBirthMonth.AddItem "August"
    cmbBirthMonth.AddItem "September"
    cmbBirthMonth.AddItem "October"
    cmbBirthMonth.AddItem "November"
    cmbBirthMonth.AddItem "December"
    
    'Fills Report Month ComboBox with Months
    cmbReportMonth.AddItem "January"
    cmbReportMonth.AddItem "February"
    cmbReportMonth.AddItem "March"
    cmbReportMonth.AddItem "April"
    cmbReportMonth.AddItem "May"
    cmbReportMonth.AddItem "June"
    cmbReportMonth.AddItem "July"
    cmbReportMonth.AddItem "August"
    cmbReportMonth.AddItem "September"
    cmbReportMonth.AddItem "October"
    cmbReportMonth.AddItem "November"
    cmbReportMonth.AddItem "December"
    
    'Fills Birth Day ComboBox with Days
    For i = 1 To 31 Step 1
        cmbBirthDay.AddItem Str(i)
    Next i
    
    'Fills Report Day ComboBox with Days
    For i = 1 To 31 Step 1
        cmbReportDay.AddItem Str(i)
    Next i
        
    'Fills Birth Year ComboBox with years from 1900 to 2100
    For i = 1900 To 2100 Step 1
        cmbBirthYear.AddItem Str(i)
    Next i
    
    'Fills Report Year ComboBox with years from 1900 to 2100
    For i = 1900 To 2100 Step 1
        cmbReportYear.AddItem Str(i)
    Next i
    
End Sub

Private Sub fillLabels()

    'Fill Center Day Label with Report Day
    lblDay(14) = DatePart("d", reportDate)
    
    'Fill Day Labels - Left Side of Chart
    n = 1
    For i = 13 To 0 Step -1
        lblDay(i) = DateAdd("d", -n, reportDate)
        lblDay(i) = DatePart("d", lblDay(i))
        n = n + 1
    Next i
    
    'Fill Day Labels - Right Side of Chart
    n = 1
    For i = 1 To 14 Step 1
        lblDay(i + 14) = DateAdd("d", n, reportDate)
        lblDay(i + 14) = DatePart("d", lblDay(i + 14))
        n = n + 1
    Next i
    
End Sub

Private Sub imgCalcBtn_Click()

    Form_Paint
    
End Sub

Private Sub lblCalcBtn_Click()

    Form_Paint
    
End Sub



Private Sub lblDay_Click(Index As Integer)

    'Move report date back to selected date
    If Index < 14 Then
        timeWarp = (14 - Index) * -1
        dateShift
    End If
    'Or move report date forward to selected date
    If Index > 14 Then
        timeWarp = Index - 14
        dateShift
    End If
    
End Sub

Private Sub userName_Change()

    'Get the User's Name
    user = userName.Text
    
End Sub

Private Sub cmbBirthMonth_Click()

    'Get the Month of Birth
    birthMonth = cmbBirthMonth.Text
    If cmbBirthMonth.Text <> "" And cmbBirthDay.Text <> "" And cmbBirthYear.Text <> "" Then
        Form_Paint
    End If
    
End Sub

Private Sub cmbBirthDay_Click()

    'Get the Day of Birth
    birthDay = cmbBirthDay.Text
        If cmbBirthMonth.Text <> "" And cmbBirthDay.Text <> "" And cmbBirthYear.Text <> "" Then
        Form_Paint
    End If
    
End Sub
Private Sub cmbBirthYear_Click()

    'Get the Year of Birth
    birthYear = cmbBirthYear.Text
    If cmbBirthMonth.Text <> "" And cmbBirthDay.Text <> "" And cmbBirthYear.Text <> "" Then
        Form_Paint
    End If

End Sub

Private Sub dateShift()

    'Implement Report Date change
    reportDate = DateAdd("d", timeWarp, reportDate)
    
    'Update Report Date Combo Boxes based on new reportDate
    reportDay = DatePart("d", reportDate)
    reportMonth = DatePart("m", reportDate)
    reportYear = DatePart("yyyy", reportDate)
    cmbReportMonth.Text = reportMonth
        If cmbReportMonth.Text = "12" Then cmbReportMonth.Text = "December"
        If cmbReportMonth.Text = "11" Then cmbReportMonth.Text = "November"
        If cmbReportMonth.Text = "10" Then cmbReportMonth.Text = "October"
        If cmbReportMonth.Text = "9" Then cmbReportMonth.Text = "September"
        If cmbReportMonth.Text = "8" Then cmbReportMonth.Text = "August"
        If cmbReportMonth.Text = "7" Then cmbReportMonth.Text = "July"
        If cmbReportMonth.Text = "6" Then cmbReportMonth.Text = "June"
        If cmbReportMonth.Text = "5" Then cmbReportMonth.Text = "May"
        If cmbReportMonth.Text = "4" Then cmbReportMonth.Text = "April"
        If cmbReportMonth.Text = "3" Then cmbReportMonth.Text = "March"
        If cmbReportMonth.Text = "2" Then cmbReportMonth.Text = "February"
        If cmbReportMonth.Text = "1" Then cmbReportMonth.Text = "January"
    cmbReportDay.Text = reportDay
    cmbReportYear.Text = reportYear
    Form_Paint

End Sub

Private Sub calc()

    Dim currentPoint As Double
    Dim bioColor As ColorConstants
    
    On Error GoTo errHandler
    
    'Clear the Screen
    bioPic.Cls
    
    'Make sure a complete BirthDate has been entered
    If cmbBirthMonth.Text = "" Or cmbBirthDay.Text = "" Or cmbBirthYear.Text = "" Then
        MsgBox "Please Enter A Valid Birth Date."
        Exit Sub
    End If
    
    birthDate = cmbBirthMonth.Text & " " & cmbBirthDay.Text & " ," & cmbBirthYear.Text
    
    'Get the Report Date
    reportDate = cmbReportMonth.Text & " " & cmbReportDay.Text & ", " & cmbReportYear.Text
    
    'Make sure reportDate is not earlier than birthDate
    If Sgn(DateDiff("d", birthDate, reportDate)) = -1 Then
        MsgBox "Report date must be later than Birth date."
        Exit Sub
    End If
    
    'Update the Day Labels
    fillLabels
    
    'Loop through biorhythm cycles and plot them
    'bioPeriod is the number of days in a cycle
    'bioColor is the color of the line to plot
    For bioNum = 1 To 4
        'Physical period
        If bioNum = 1 Then bioPeriod = 23: bioColor = vbRed
        'Emotional period
        If bioNum = 2 Then bioPeriod = 28: bioColor = vbBlue
        'Intellectual period
        If bioNum = 3 Then bioPeriod = 33: bioColor = vbGreen
        'Intuitional period
        If bioNum = 4 Then bioPeriod = 38: bioColor = vbCyan

        'Find the first peak to left of middle
        currentPoint = bioPic.ScaleWidth / 2 - ((DateDiff("d", birthDate, reportDate) Mod bioPeriod) - bioPeriod / 4)
        'Find the first peak that is off the chart
        Do While currentPoint > 0
            currentPoint = currentPoint - bioPeriod
        Loop
        
        'Necessary because next loop adds BioPeriod/2 back to the variable
        currentPoint = currentPoint - bioPeriod / 2
        
        'Find high and low points and plot parabolas
        Do While currentPoint < bioPic.ScaleWidth
            currentPoint = currentPoint + bioPeriod / 2
                 If currentPoint + bioPeriod / 4 >= 0 Then
                Parabola bioNum, currentPoint, 0, currentPoint + bioPeriod / 4, bioColor
            End If
                currentPoint = currentPoint + bioPeriod / 2
            If currentPoint + bioPeriod / 4 >= 0 Then
                Parabola bioNum, currentPoint, bioPic.ScaleHeight, currentPoint + bioPeriod / 4, bioColor
            End If
        Loop
    
    Next bioNum
    
errHandler:
    If Err.Number <> 0 Then
        MsgBox "Birth Date or Report Date is incorrect."
        Exit Sub
    End If
    
End Sub

Public Function Parabola(bioNum As Integer, Xa As Double, Ya As Double, lastPt As Double, RedGreenBlueCyan As ColorConstants)
    
    'This function creates a parabola when vertex and last point are given
    'Vertex = Xa, Ya
    'LastPoint = LastPt, horCenter
    Dim horCenter As Integer
    Dim slope As Double
    Dim Y As Double, X As Double
        
        'Set horCenter to be the horizontal center of the biorhythm picture
        horCenter = bioPic.ScaleHeight / 2
        'Find slope of parabola
        slope = (horCenter - Ya) / ((lastPt - Xa) ^ 2)
        'Graph the parabola
        For X = (Xa - (lastPt - Xa)) To lastPt Step 0.01
            Y = slope * ((X - Xa) ^ 2) + Ya
            bioPic.PSet (X, Y), RedGreenBlueCyan
        Next X

End Function

