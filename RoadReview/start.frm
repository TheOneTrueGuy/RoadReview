VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form start 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Job assignment"
   ClientHeight    =   5235
   ClientLeft      =   1785
   ClientTop       =   4365
   ClientWidth     =   7410
   Icon            =   "start.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Load from disk"
      Default         =   -1  'True
      Height          =   390
      Left            =   5520
      TabIndex        =   27
      Top             =   3900
      Width           =   1620
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create new job"
      Height          =   390
      Left            =   5520
      TabIndex        =   9
      Top             =   4350
      Width           =   1620
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      TabIndex        =   26
      Top             =   4800
      Width           =   1620
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1065
      TabIndex        =   0
      Top             =   45
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Set job as default"
      Height          =   225
      Left            =   4185
      TabIndex        =   23
      ToolTipText     =   "This job wil be loaded at startup (skip start screen)"
      Top             =   60
      Width           =   1605
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2850
      TabIndex        =   22
      Top             =   3840
      Width           =   810
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H007FFF7F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   30
      ScaleHeight     =   1635
      ScaleWidth      =   7365
      TabIndex        =   13
      ToolTipText     =   "Values relating to base plane and benchmark"
      Top             =   390
      Width           =   7365
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   5010
         TabIndex        =   3
         Top             =   330
         Width           =   600
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3300
         TabIndex        =   2
         Top             =   330
         Width           =   585
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1695
         TabIndex        =   1
         Top             =   38
         Width           =   930
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1245
         Width           =   825
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   825
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1770
         TabIndex        =   4
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Station Row"
         Height          =   225
         Left            =   4050
         TabIndex        =   21
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Station Column"
         Height          =   210
         Left            =   2175
         TabIndex        =   20
         Top             =   367
         Width           =   1245
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Bench mark station location:"
         Height          =   210
         Left            =   0
         TabIndex        =   19
         Top             =   367
         Width           =   2085
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bench Mark elevation:"
         Height          =   240
         Left            =   0
         TabIndex        =   18
         Top             =   60
         Width           =   1710
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Slope along length (if known):"
         Height          =   255
         Left            =   15
         TabIndex        =   17
         Top             =   1305
         Width           =   2160
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Slope along width (if known):"
         Height          =   225
         Left            =   15
         TabIndex        =   16
         Top             =   1005
         Width           =   2115
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Bench mark elevation features"
         Height          =   240
         Left            =   5160
         TabIndex        =   15
         Top             =   0
         Width           =   2310
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual plane elevation"
         Height          =   240
         Left            =   30
         TabIndex        =   14
         Top             =   660
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   885
      ScaleWidth      =   7380
      TabIndex        =   10
      ToolTipText     =   "Values relating to current elevations"
      Top             =   2040
      Width           =   7380
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2685
         TabIndex        =   8
         Text            =   "25"
         Top             =   435
         Width           =   630
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1485
         TabIndex        =   7
         Top             =   30
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Station spacing distance (in feet):"
         Height          =   270
         Left            =   30
         TabIndex        =   12
         ToolTipText     =   "Equal distance grids needs only one set"
         Top             =   465
         Width           =   2430
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of stations:"
         Height          =   210
         Left            =   15
         TabIndex        =   11
         Top             =   75
         Width           =   1485
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6330
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "lfd"
   End
   Begin VB.Label Label1 
      Caption         =   "Job Name"
      Height          =   180
      Left            =   75
      TabIndex        =   25
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label8 
      Caption         =   "Expected compaction rate in percents:"
      Height          =   285
      Left            =   60
      TabIndex        =   24
      Top             =   3885
      Width           =   2820
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stations() As station, proposed_elev() As station
Public loaded_from_disk As Boolean, set_default As Boolean
Public landname As String, gnu As Boolean, opn_gnu As Boolean
Public numx As Integer, numy As Integer, standard As Double
Public benchx As Integer, benchy As Integer
Public wstation_distance As Double, lstation_distance As Double
Public wslope As Double, lslope As Double, factor As Integer
Public cam_height As Integer, rotate_angle As Integer, escale As Integer
Public UpDownZoom As Integer, HScroll1 As Integer, VScroll1 As Integer
Public xcen2 As Double, ycen2 As Double, screen_dist As Double

Private Sub Check1_Click()
' make sure default passed to props
If Check1.Value = 1 Then set_default = True
End Sub
Private Sub Command1_Click()
CommonDialog1.DefaultExt = "vee"
CommonDialog1.Filter = "Volume estimation files (*.vee)|*.vee"
CommonDialog1.CancelError = True
On Error GoTo erhandl
again:
CommonDialog1.ShowOpen
If opn_gnu Then Unload VMain
landname$ = CommonDialog1.FileName
If landname = "" Then MsgBox ("You must select a file"): GoTo again:
'math.init_mod
start.Visible = False
VMain.Show: VMain.SetFocus
'If VMain.aEditor.editing Then Call VMain.view_editor_Click
erhandl:
End Sub

Private Sub Command2_Click()
 If gnu Then Unload VMain
If Check1.Value = 1 Then
 Open "vee.ini" For Output As #1
 Print #1, Text1.Text
 Close 1
End If
If Val(Text2.Text) < 2 Then Text2.Text = "2"
If Val(Text3.Text) < 2 Then Text3.Text = "2"
math.loaded_from_disk = False
VMain.loaded_from_disk = False
'math.init_mod
VMain.Show: VMain.SetFocus
math.data_changed = True
started = True
End Sub

Private Sub Command3_Click()
If gnu Or opn_gnu Then start.Visible = False: Exit Sub
End
End Sub

Private Sub Form_Load()
copyright$ = "copyright Guy Giesbrecht 1999"
On Error GoTo noini
Open "vee.ini" For Input As #1
Input #1, ini$
Close 1

VMain.Show: VMain.SetFocus
start.Visible = False
Exit Sub

noini: ' or a disk error
End Sub

Public Sub gnuopn()
Command1_Click
End Sub

Private Sub Option1_Click(index As Integer)
factor = index
End Sub
