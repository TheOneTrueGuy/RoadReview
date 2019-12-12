VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Data_Entry2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Station Data"
   ClientHeight    =   7545
   ClientLeft      =   915
   ClientTop       =   1125
   ClientWidth     =   10980
   Icon            =   "Data_Entry2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10980
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   9075
      TabIndex        =   14
      Top             =   2325
      Width           =   810
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   9060
      TabIndex        =   13
      Top             =   1920
      Width           =   840
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Begin Taper in This Cell"
      Height          =   405
      Left            =   7965
      TabIndex        =   12
      Top             =   3960
      Width           =   2670
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start curve in current cell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7965
      TabIndex        =   11
      Top             =   3345
      Width           =   2805
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9045
      TabIndex        =   10
      Text            =   "8.5"
      Top             =   1545
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9060
      TabIndex        =   8
      Text            =   "8"
      Top             =   1215
      Width           =   840
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9810
      TabIndex        =   6
      Text            =   "1000"
      Top             =   2985
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10110
      TabIndex        =   4
      Text            =   "25"
      Top             =   2670
      Width           =   705
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10215
      Top             =   570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Road View"
      Height          =   525
      Left            =   7965
      TabIndex        =   2
      Top             =   630
      Width           =   1170
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9540
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   90
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Data_Entry2.frx":030A
      Height          =   315
      Left            =   8220
      OleObjectBlob   =   "Data_Entry2.frx":031E
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7410
      Left            =   -15
      TabIndex        =   0
      Top             =   30
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   13070
      _Version        =   393216
      Rows            =   100
      Cols            =   6
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   3
      FormatString    =   "Station    | Lat                     |   Long                | Elevation  | Width Left | Width Right"
   End
   Begin VB.Label Label5 
      Caption         =   "Slope Right"
      Height          =   300
      Index           =   1
      Left            =   7980
      TabIndex        =   16
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label Label5 
      Caption         =   "Slope Left"
      Height          =   300
      Index           =   0
      Left            =   7980
      TabIndex        =   15
      Top             =   1935
      Width           =   1050
   End
   Begin VB.Label Label4 
      Caption         =   "Width right"
      Height          =   270
      Left            =   7980
      TabIndex        =   9
      Top             =   1545
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Width Left"
      Height          =   270
      Left            =   7980
      TabIndex        =   7
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of stations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7965
      TabIndex        =   5
      Top             =   3015
      Width           =   1830
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Station Distance in feet:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7965
      TabIndex        =   3
      Top             =   2685
      Width           =   2130
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu opn 
         Caption         =   "Open"
      End
      Begin VB.Menu import 
         Caption         =   "Import Excel"
      End
      Begin VB.Menu xport 
         Caption         =   "Export database"
      End
      Begin VB.Menu sav 
         Caption         =   "Save"
      End
      Begin VB.Menu Xit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Data_Entry2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loading As Boolean, loaded As Boolean, clickerase As Boolean
Const cpyright$ = "Copyright Guy Giesbrecht 1999"
Public copyright$
Dim cellcolor As Long, oldbcolor As Long
Dim xcell As Double, ycell As Double
Dim benchmark As Poynt

'' ___________________________
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim dato As ADODB.Recordset
Dim datcon As ADODB.Connection
Dim dacomm As ADODB.Command
Dim strConn$


Dim xlName As String
Dim SheetName As String
Dim mousebutton As Integer
Dim mousekey As Integer
Dim oldcol As Long, oldrow As Long
Dim stationdistance As Double, numstations As Long
Dim datar() As Double
Dim latar() As Double, lonar() As Double
Dim elar() As Double
'Dim curves(), curvecount As Long, curveopen As Boolean
Dim origin As Poynt
Dim roadpiece() As station
Dim cellind

' sationdistance sets fundamental scale.
Public Sub Segmentor()
' calculate road segments for 6 points:
' road center1, road center2
' pointleft=roadcenter-widthleft
' pointright=roadcenter+widthright

Dim tzl, seg As segment
For tzl = 1 To UBound(datar, 2) Step 2

Next tzl
' first point needs to be centered at 0 so that successive point
' elevations are p2-p1



End Sub
Private Sub Command1_Click()
loading = True
Dim xlz, zlz, ary(3, 10) As Double
For xlz = 0 To 2
For zlz = 1 To 10
ary(xlz, zlz) = Rnd * 2
Next zlz
Next xlz


RoadView.Show

Exit Sub
Grid1.Row = 1: Grid1.Col = 1
benchmark.x = CSng(Grid1.Text)
Grid1.Col = 2
benchmark.z = CSng(Grid1.Text)
Grid1.Col = 3
benchmark.y = CSng(Grid1.Text)

Dim lat As Double, longi As Double, elev As Double
Dim yzl, lft As Double, rit As Double

For yzl = 1 To 10
'Grid1.Row = yzl
'Grid1.Col = 1

lat = datar(0, yzl) 'CSng(Grid1.Text)
'Grid1.Col = 2

longi = datar(1, yzl) ' CSng(Grid1.Text)
'Grid1.Col = 3

elev = datar(2, yzl) 'CSng(Grid1.Text)
lft = datar(2, yzl) - datar(3, yzl)
rit = datar(2, yzl) + datar(4, yzl)
'RoadView.loadVertice lat, longi, elev


Next yzl
'RoadView.Show

'RoadView.Render
End Sub

Private Sub Command2_Click()
'If curveopen Then MsgBox ("Previous curve not yet closed"): Exit Sub
'math.curves(math.curvecount).row = Grid1.row

cellind = Grid1.Row
curveStats.Show
End Sub
Public Sub tranfer_curve()
loading = True
Dim pnts() As Poynt, tzl
pnts = crvRange(math.curve1)

For tzl = 0 To UBound(pnts)

' add points into appropriate cells in grid1
Grid1.Col = 1
If cellind + tzl > Grid1.rows - 1 Then
Grid1.rows = cellind + tzl + 1
ReDim Preserve datar(Grid1.Cols, tzl + 1)
End If
datar(0, tzl) = pnts(tzl).x

Grid1.Row = cellind + tzl
Grid1.Text = Cdegrees(pnts(tzl).x)
Grid1.Col = 2
datar(1, tzl) = pnts(tzl).z
Grid1.Text = Cdegrees(pnts(tzl).z)
Grid1.Col = 0
Grid1.Row = cellind + tzl
Grid1.Text = cellind + tzl
Next tzl
loading = False
End Sub
Public Function getDatar() As Double()
'
getDatar = datar

End Function

Private Sub Command3_Click()
loadupDB
'If curves(0, curvecount) < Grid1.Row Then MsgBox ("Curve End must follow curve start"): Exit Sub
'curves(1, curvecount) = Grid1.Row
'curvecount = curvecount + 1

End Sub

Private Sub Form_Activate()
Grid1_EnterCell
End Sub

Private Sub Form_GotFocus()
' run data reloader
'load_data
End Sub

Private Sub Form_Deactivate()
math.data_changed = data_changed
Grid1_LeaveCell
End Sub

Private Sub Form_Load()
ReDim datar(Grid1.Cols, Grid1.rows)
ReDim latar(40), lonar(40), elar(40)
ReDim curves(2, 3)
loading = True
standard = math.standard
Dim xc As Integer, yr As Integer, xk As Integer

    numx = math.numx * 2 + 1: numy = math.numy + 1
    benchx = math.benchx / 2: benchy = math.benchy
    'ReDim numexcep(numx, numy)
    copyright$ = "copyright Guy Giesbrecht jan 1, 1999"
    Grid1.Col = 0: Grid1.Row = 0
'    Grid1.Cols = numx
'    Grid1.Rows = numy
    Grid1.Row = 0
    
    
    Grid1.Col = 0
    For yr = 0 To numy - 1
      Grid1.Row = yr: Grid1.Text = CStr(yr)
    Next yr
    Dim cl1 As Integer, cl2 As Integer, xx As Integer, yy As Integer
    
    'If math.loaded_from_disk Then load_data

    For xc = 1 To numx - 1
        Grid1.Col = xc
        For yr = 1 To numy - 1
        Grid1.Row = yr
        If xc / 2 = Int(xc / 2) Then Grid1.CellBackColor = RGB(127, 127, 255) Else Grid1.CellBackColor = RGB(255, 127, 127)
        Next yr
    Next xc

    Grid1.Col = 0: Grid1.Row = 0
        Grid1.CellBackColor = RGB(127, 255, 127)
    cellcolor = Grid1.CellBackColor
   stationdistance = 25 ' feet that is
   cellind = 1
loading = False
End Sub
'Public Sub load_data()
'Dim cl1 As Integer, cl2 As Integer, xx As Integer, yy As Integer
'For xx = 1 To numx - 1
'        If xx / 2 <> Int(xx / 2) Then cl1 = cl1 + 1 Else cl2 = cl2 + 1
'        Grid1.Col = xx
'        For yy = 1 To numy - 1
'         Grid1.Row = yy
'        If math.elev(cl1, yy) < 0 Then GoTo skipper
'         If xx / 2 <> Int(xx / 2) Then
'         Grid1.Text = Str(math.elev(cl1, yy))
'         Else:
'          Grid1.Text = Str(math.prop_elev(cl2, yy))
'         End If
'skipper:
'        Next yy
'        Next xx
'
'End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
xcell = x: ycell = y
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Grid1_Click()
' setup double click cell erase on backspace
Debug.Print Grid1.Col, Grid1.Row
clickerase = False
End Sub
Private Sub Grid1_DblClick()
clickerase = True
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
data_changed = True
If KeyAscii = 13 Then Grid1_LeaveCell: Exit Sub
If clickerase Then
   Grid1.Text = "": clickerase = False:
   If KeyAscii = 8 Then Exit Sub
End If
If KeyAscii = 8 And Len(Grid1.Text) > 0 Then Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1): Exit Sub

Grid1.Text = Grid1.Text + Chr(KeyAscii)

Dim yin As Integer

    yin = Grid1.Row
   
End Sub

Private Sub Grid1_EnterCell()
If loading Then Exit Sub
'If Grid1.Col = oldcol And Grid1.Row = oldrow Then Exit Sub ' no new selection

cellcolor = Grid1.CellBackColor
Grid1.CellBackColor = RGB(255, 255, 255)
Debug.Print "EC:"; Grid1.Col, Grid1.Row

End Sub

Public Sub getExcep_pointer()
'Dim temElev As station
'' put x,y and excep# in globals
'For xn = -1 To 1
'For yn = -1 To 1
'
'If grx + xn < 1 Then GoTo skipc
'If grx + xn > numx - 1 Then GoTo skipc
'If gry + yn < 1 Then GoTo skipc
'If gry + yn > numy - 1 Then GoTo skipc
'If curorprop Then temElev = math.get_elev(grx + xn, gry + yn) Else temElev = math.get_prop_elev(grx + xn, gry + yn)
'' convert from feature(n,n) to string
'' extract pointer from ListIndex
'For nx = 0 To temElev.num_excep
'zeep = zeep + 1
'If zeep = List2.ListIndex Then Exit For
'Next nx
'skipc:
'If zeep = List2.ListIndex Then Exit For
'Next yn
'If zeep = List2.ListIndex Then Exit For
'Next xn
''
'skipexcept:
'' don't forget to increment nex_exc_counter  properly.
'' In calling routine: preset to 0
'nex_exc_pointer(0, nex_exc_counter) = xn: nex_exc_pointer(1, nex_exc_counter) = yn
'nex_exc_pointer(2, nex_exc_counter) = nx

End Sub


Private Sub Grid1_LeaveCell()
If loading Then Exit Sub
Dim curcol As Long, currow As Long
Grid1.CellBackColor = cellcolor
Debug.Print "LC:"; Grid1.Col, Grid1.Row

If Grid1.Row = 1 Then
' set banchmark
End If

Select Case Grid1.Col
Case 1
'latar(currow) = parseDegrees(Grid1.Text)
Case 2
'lonar(currow) = parseDegrees(Grid1.Text)
Case 3
elar(currow) = CSng(Grid1.Text)

Case 4
Case 5
Case 6
End Select

'If curcol = 1 Then
'
'End If
'If curcol = 2 Then
'
'End If

oldcol = Grid1.Col: oldrow = Grid1.Row ' must be last line in sub
If Grid1.Col = 3 Then
If Grid1.Text <> "" Then datar(Grid1.Col, Grid1.Row) = CSng(Grid1.Text) ' background data array
End If
' for use so that unfilled cols and rows don't through errors in
' the vertice/image generation routines.

End Sub

Public Function loadupDB() As Boolean
'On Error GoTo ercl
Dim rows
ChDir App.Path
Set dato = CreateObject("ADODB.Recordset")
Set datcon = CreateObject("ADODB.Connection")
Set dacomm = CreateObject("ADODB.Command")
strConn$ = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = db1.mdb"
datcon.Open strConn
dato.Open "Table1", datcon

'Do While Not dato.EOF
'
'rows = rows + 1
'Loop
loadupDB = True
Exit Function
ercl:
loadupDB = False
End Function



Public Sub importExcel(dbName$)
Dim dbnumx As Integer, dbnumy As Integer
'CreateRS dbName$
dbnumx = DBGrid1.Columns.Count
Dim z
For z = 1 To dbnumx
DBGrid1.Col = z
If DBGrid1.Text = "" Then Exit For
Next
dbnumx = z
dbnumy = rs.RecordCount
math.reInit dbnumx, dbnumy
Grid1.Cols = dbnumx
Grid1.rows = dbnumy
math.numx = dbnumx - 1
math.numy = dbnumy - 1
Dim xi, yi

For xi = 1 To dbnumx - 1
Grid1.Col = xi
DBGrid1.Col = xi
For yi = 1 To dbnumy - 1
Grid1.Row = yi
DBGrid1.Row = yi
Grid1.Text = DBGrid1.Text
Next yi
Next xi
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "MD"; Grid1.Col, Grid1.Row
mousebutton = Button
Debug.Print "mdb"; mousebutton
mousekey = Shift
If mousebutton = 2 Then
' transfer values between first and selected rows
End If
End Sub

Private Sub Grid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' if right click then  pop up tiling, transform, other data control methods
End Sub

Private Sub Grid1_SelChange()
Debug.Print Grid1.Col, Grid1.Row
If Grid1.Col = 0 Then
' clicked on 0 col, setup for context menu
' begin curve, end curve etc.

End If

End Sub

Private Sub import_Click()
'
Dim fyl$, filt$

CommonDialog1.Filter = "Excel | *.xls; *.mdb"
CommonDialog1.ShowOpen
fyl = CommonDialog1.FileName
If fyl = "" Then Exit Sub

CommonDialog1.Filter = ""

End Sub

Private Sub opn_Click()
'
Dim fyl$
CommonDialog1.ShowOpen
fyl = CommonDialog1.FileName
If fyl = "" Then Exit Sub


End Sub

Sub CreateDatabase()

   Dim cat As New ADOX.Catalog
   
   cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\new.mdb"

End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbCrLf Then
' enter new value
stationdistance = CSng(Text1)
End If
End Sub

Private Sub Text1_LostFocus()
'
' enter new value into
stationdistance = CSng(Text1)

End Sub

Private Sub Text2_LostFocus()
numstations = CLng(Text2)
End Sub




Private Function crvRange(Crv As Curve) As Poynt()
' D should increment by d1 each succesive call
' calculate curve values
' C=2*R*sin(D/2) ' chord length for angle D
' T=R*tan(A/2) ' here A signifies delta
' LC = 2*R*Sin(A/2) ' Length of long chord, for angle delta (A)
'
' Remember that coords need to be offset by R from center of road (to make
' center of circle with Radius R be correct)

' create deflection table
Dim LC As Double
LC = 2 * Crv.R * Sin(Crv.A / 2)

Dim deftab() As Double, acdef As Double, defpts
'If Crv.leftright Then
' curves left
'Else
' curves right

Dim x, z, d2, pnt() As Poynt, xa, za
Dim numpnts As Double, pc As Double

numpnts = Crv.L / stationdistance

While defpts < Int(numpnts) ' acdef <Crv.A / 2
    acdef = acdef + Crv.D / 2
    ReDim Preserve deftab(defpts)
    deftab(defpts) = acdef
    defpts = defpts + 1
    If acdef + Crv.D / 2 > Crv.A / 2 And acdef <> Crv.A / 2 Then
    ReDim Preserve deftab(defpts)
    deftab(defpts) = Abs(Crv.A / 2 - acdef)
    acdef = acdef + Crv.D
    End If
Wend
'End If


' numpts and defpts must agree !!!

ReDim pnt(numpnts)
Dim tzl, ra, na, xo, zo
na = 0
'pc = Crv.PCn
xa = 0 'Crv.PCn.x
za = 0 'Crv.PCn.z
ra = 2 * Crv.R * Sin(deftab(na))
'na = na + 1 'tan_angle(xa, za)
'd2 = Crv.D
d2 = deftab(na) 'na + d2
Dim xs, zc
For tzl = 1 To numpnts
xs = Sin(d2) * ra
zc = Cos(d2) * ra
x = xo + xs ' offset from previous set
z = zo + zc
If x > 360 Or z > 360 Then Stop
'ra = Sqr(xa ^ 2 + ya ^ 2)
'xo = x: zo = z
na = na + 1
d2 = deftab(na) 'd2 + Crv.D
pnt(tzl).x = x
pnt(tzl).z = z
Next tzl
' still needs to calc final EC point chord!!
crvRange = pnt

End Function

Private Function curveRange(Crv As Curve) As Poynt()
' D should increment by d1 each succesive call
' calculate curve values
' C=2*R*sin(D/2) ' chord length for angle D
' T=R*tan(A/2) ' here A signifies delta
' LC = @*R*Sin(A/2) ' Length of long chord, for angle delta (A)
'
' Remember that coords need to be offset by R from center of road (to make
' center of circle with Radius R be correct)
Dim x, z, d2, pnt() As Poynt
Dim numpnts As Double
numpnts = Crv.L / stationdistance
ReDim pnt(numpnts)
Dim tzl
d2 = Crv.D
For tzl = 1 To numpnts
x = Sin(d2) * Crv.R
z = Cos(d2) * Crv.R
d2 = d2 + Crv.D
pnt(tzl).x = x
pnt(tzl).z = z
Next tzl
curveRange = pnt

End Function

Public Function transformcoordAxis()
' road co-ords not lined up along a standard axis
' must have their values transformed to align curves with
' the correct movement
' calculate "slope" of road in x-z plane
' use right angle extension of R from PC and slope value
' to calculate offset
' due to tangent rule, offset need only be the dX and dY of PT
' and first point calculated by curveRange(), all succeeding points
' will match the same offset
' note that quadrant of rotation is critical
End Function



Public Sub convertToLatLonElev(dbName$)
' take the x,y,z offset values from the data array
' and add or subtract from the origin coords (lat+x,lon+z,elev+y) in degrees,
' minutes, and seconds. Save as database


'####_____ This stuff is just the db stuff
'#  The DB is formatted: Lat, Elev, Lon
Dim rows
ChDir App.Path
Set dato = CreateObject("ADODB.Recordset")
Set datcon = CreateObject("ADODB.Connection")
Set dacomm = CreateObject("ADODB.Command")
strConn$ = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = db1.mdb"
datcon.Open strConn
dato.Open "Table1", datcon

End Sub

Private Sub Xit_Click()
Unload Me
End
End Sub

Private Sub xport_Click()
'
Dim fyl$, filt$

CommonDialog1.Filter = "Roadview database(*.mdb)| *.mdb"
CommonDialog1.ShowSave
fyl = CommonDialog1.FileName
If fyl = "" Then Exit Sub
'_______________]
Dim cat As New ADOX.Catalog
cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & fyl

Dim rows
Dim dat1 As ADODB.Recordset
Dim datc1 As ADODB.Connection
Dim datcm As ADODB.Command
ChDir App.Path
Set dat1 = CreateObject("ADODB.Recordset")
Set datc1 = CreateObject("ADODB.Connection")
Set datcm = CreateObject("ADODB.Command")
strConn$ = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & fyl
datc1.Open strConn
dat1.Open "Table1", datc1

'Do While Not dat1.EOF
'
'rows = rows + 1
'Loop
'____________________________________

CommonDialog1.Filter = ""

End Sub
