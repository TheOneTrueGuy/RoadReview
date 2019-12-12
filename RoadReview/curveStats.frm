VERSION 5.00
Begin VB.Form curveStats 
   Caption         =   "Curve Ranging Data"
   ClientHeight    =   5865
   ClientLeft      =   4380
   ClientTop       =   1875
   ClientWidth     =   7365
   Icon            =   "curveStats.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7365
   Begin VB.TextBox Text8 
      Height          =   285
      Index           =   1
      Left            =   6405
      TabIndex        =   34
      Top             =   3660
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Index           =   0
      Left            =   4620
      TabIndex        =   32
      Top             =   3675
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   330
      Left            =   2955
      TabIndex        =   12
      Text            =   "6d00m00"
      Top             =   3000
      Width           =   1155
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2970
      TabIndex        =   10
      Text            =   "4d00m00"
      Top             =   2610
      Width           =   1140
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2970
      TabIndex        =   8
      Text            =   "2d00m00"
      Top             =   2205
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1065
      TabIndex        =   11
      Text            =   "5d00m00"
      Top             =   2985
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1065
      TabIndex        =   9
      Text            =   "3d00m00"
      Top             =   2610
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1050
      TabIndex        =   7
      Text            =   "1d00m00"
      Top             =   2220
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3330
      TabIndex        =   13
      Text            =   ".037"
      Top             =   210
      Width           =   1260
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3330
      TabIndex        =   14
      Text            =   "+1"
      Top             =   630
      Width           =   1275
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3330
      TabIndex        =   15
      Text            =   "1"
      Top             =   1050
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Curves Right"
      Height          =   285
      Index           =   1
      Left            =   2055
      TabIndex        =   1
      Top             =   1755
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Curves Left"
      Height          =   285
      Index           =   0
      Left            =   2055
      TabIndex        =   0
      Top             =   1455
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   1020
      TabIndex        =   17
      Top             =   3750
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   450
      Left            =   165
      TabIndex        =   16
      Top             =   3750
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   405
      TabIndex        =   6
      Text            =   "4583.66"
      Top             =   1830
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   405
      TabIndex        =   5
      Text            =   "4276.8"
      Top             =   1425
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   405
      TabIndex        =   4
      Text            =   "2308.35"
      Top             =   1020
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   405
      TabIndex        =   3
      Text            =   "1d15m00"
      Top             =   630
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   405
      TabIndex        =   2
      Text            =   "53d27m36"
      Top             =   225
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "Longitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   5295
      TabIndex        =   33
      Top             =   3660
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "Origin Point Latitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   2430
      TabIndex        =   31
      Top             =   3690
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Long"
      Height          =   210
      Index           =   1
      Left            =   2415
      TabIndex        =   30
      Top             =   2250
      Width           =   390
   End
   Begin VB.Label Label4 
      Caption         =   "Lat"
      Height          =   210
      Index           =   0
      Left            =   660
      TabIndex        =   29
      Top             =   2250
      Width           =   315
   End
   Begin VB.Label lab 
      Caption         =   "PT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   30
      TabIndex        =   28
      Top             =   2970
      Width           =   510
   End
   Begin VB.Label lab 
      Caption         =   "PI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   45
      TabIndex        =   27
      Top             =   2565
      Width           =   510
   End
   Begin VB.Label lab 
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   60
      TabIndex        =   26
      Top             =   2175
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Super Elevation:"
      Height          =   285
      Left            =   1935
      TabIndex        =   25
      Top             =   225
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Width Left"
      Height          =   285
      Left            =   2370
      TabIndex        =   24
      Top             =   630
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "Width Right"
      Height          =   285
      Left            =   2340
      TabIndex        =   23
      Top             =   1065
      Width           =   930
   End
   Begin VB.Label lab 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   105
      TabIndex        =   22
      Top             =   1815
      Width           =   285
   End
   Begin VB.Label lab 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   105
      TabIndex        =   21
      Top             =   1398
      Width           =   285
   End
   Begin VB.Label lab 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   105
      TabIndex        =   20
      Top             =   982
      Width           =   285
   End
   Begin VB.Label lab 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   105
      TabIndex        =   19
      Top             =   566
      Width           =   285
   End
   Begin VB.Label lab 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   105
      TabIndex        =   18
      Top             =   150
      Width           =   285
   End
End
Attribute VB_Name = "curveStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' C=2*R*sin(D/2) ' chord length for angle D
' T=R*tan(A/2) ' here A signifies delta
' LC = @*R*Sin(A/2) ' Length of long chord, for angle delta (A)

Public A1 As Double, D As Double, T As Double, L As Double, R As Double, PCn As Double
Dim Crange() As Poynt


Private Sub Command1_Click()
A1 = parseDegrees(Text1(0))
D = parseDegrees(Text1(1))
T = CDbl(Text1(2))
L = CSng(Text1(3))
R = CDbl(Text1(4))
'PCn = parseDegrees(Text1(5)) ' PC in DMS
Dim curly As Curve
curly.A = A1
curly.D = D
curly.L = L
curly.T = T
curly.R = R
curly.PCn.x = parseDegrees(Text1(5))
curly.PCn.z = parseDegrees(Text5)
curly.leftright = Option1(0).Value
math.curve1 = curly
ReDim math.curves(math.curvecount)
math.curves(math.curvecount) = curly
math.curvecount = math.curvecount + 1
'Crange = curveRange
Data_Entry2.tranfer_curve
Me.Hide
End Sub
Public Function getCurvePointCount(index As Long) As Double
'
End Function
Public Function getCurvePointX(index As Long) As Double
'
End Function
Public Function getCurvePointY(index As Long) As Double
'
End Function

Private Function curveRange() As Poynt()
' D should increment by d1 each succesive call
' calculate curve values
' C=2*R*sin(D/2) ' chord length for angle D
' T=R*tan(A/2) ' here A signifies delta
' LC = @*R*Sin(A/2) ' Length of long chord, for angle delta (A)
'
' Remember that coords need to be offset by R from center of road (to make
' center of circle with Radius R be correct)

Dim x, y, d2, pnt() As Poynt
Dim numpnts As Double
numpnts = Lcurve / stationdistance
ReDim pnt(numpnts)
Dim tzl
d2 = D
For tzl = 1 To numpnts
x = Sin(d2) * R
y = Cos(d2) * R
d2 = d2 + D
pnt(tzl).x = x
pnt(tzl).y = y
Next tzl
curveRange = pnt

End Function

