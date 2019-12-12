Attribute VB_Name = "Math2a"
Dim elevations() As Elevation, proposed_elev() As Elevation, benchplane() As Elevation
Dim elevon() As Elevation
Public wslope As Single, lslope As Single
Public wdist As Single, ldist As Single
Public standard As Single, shrink As Single
Public offset As Single, resoscale As Single
Dim baseline As Single, baseline2 As Single
Dim numx As Integer, numy As Integer
Public cut_vol As Single, fill_vol As Single
Public pcut_vol As Single, pfill_vol As Single
Public dcut_vol As Single, dfill_vol As Single
Dim popeye As Boolean
Public Sub init_mod()
' get all the stuff from math initial to type of elevation being estimated
wdist = math.wstation_distance: ldist = math.lstation_distance
wslope = math.wslope: lslope = math.lslope
standard = math.standard
numx = math.numx: numy = math.numy
ReDim elevations(numx, numy), proposed_elev(numx, numy)
ReDim elevon(numx, numy)
Dim xx As Integer, yy As Integer
For xx = 1 To numx
 For yy = 1 To numy
    elevations(xx, yy) = math.get_elev(xx, yy)
    proposed_elev(xx, yy) = math.get_prop_elev(xx, yy)
 Next yy
 Next xx
If shrink = 0 Then shrink = 1
resoscale = 400
End Sub
Public Function total_cut_vol(typof As Integer) As Single
'Printer.Print "total_cut_vol"
Dim xx As Integer, yy As Integer, total_cvol As Single
init_mod
If typof = 1 Then load_current
If typof = 2 Then load_pro
For xx = 1 To numx - 1
  For yy = 1 To numy - 1
    total_cvol = total_cvol + calc_cube(xx, yy, True)
  Next yy
Next xx
'Printer.Print "total_cvol: "; total_cvol / resoscale
total_cut_vol = total_cvol / resoscale ' make this value resolution (resoscale) scale later
End Function
Public Function total_fill_vol(typof As Integer) As Single
'Printer.Print "total_fill_vol"
init_mod
Dim xx As Integer, yy As Integer, total_fvol As Single
If typof = 1 Then load_current
If typof = 2 Then load_pro
For xx = 1 To numx - 1
  For yy = 1 To numy - 1
    total_fvol = total_fvol + calc_cube(xx, yy, False)
  Next yy
Next xx
'Printer.Print "total_fvol: "; total_fvol / resoscale
total_fill_vol = total_fvol / resoscale ' make this value resolution scale (resoscale) later
End Function

Public Sub load_current()
'debug.Print "Current"
For xx = 1 To numx
  For yy = 1 To numy
  elevon(xx, yy) = elevations(xx, yy)
'  debug.Print elevon(xx, yy).elev(0), elevon(xx, yy).elev(1), elevon(xx, yy).elev(2)
  Next yy
Next xx
End Sub
Public Sub load_pro()
'debug.Print "proposed"
For xx = 1 To numx
  For yy = 1 To numy
  elevon(xx, yy) = proposed_elev(xx, yy)
'   debug.Print elevon(xx, yy).elev(0), elevon(xx, yy).elev(1), elevon(xx, yy).elev(2)
  Next yy
Next xx
End Sub
Public Function calc_cube(xindex As Integer, yindex As Integer, cutorfill As Boolean) As Single
'Printer.Print "calc_cube: "; xindex; ","; yindex; ","; cutorfill
' Big hang up now: averaging of sides which show no exceptions
Dim curelev As Elevation, mdflag As Boolean
Dim nexelevx As Elevation, nexelevy As Elevation, nexelevxy As Elevation
Dim excepsx As Elevation, excepsy As Elevation
Dim excepsfront As Elevation, excepsright As Elevation
Dim mdexcepsx As Elevation, mdexcepsy As Elevation
curelev = elevon(xindex, yindex)
nexelevy = elevon(xindex, yindex + 1)
nexelevx = elevon(xindex + 1, yindex)
nexelevxy = elevon(xindex + 1, yindex + 1)
Dim exc As Integer, excy As Integer, excf As Integer, excr As Integer
' must put in subelevs for beginning and end elevs on mdexceps

' find all the exceps on first face --------------------------------------------
For testexcepsy = 0 To curelev.num_excep
    If curelev.xdist(testexcepsy) = 0 Then
    excepsy.xdist(excy) = curelev.ydist(testexcepsy) 'y
    excepsy.elev(excy) = curelev.elev(testexcepsy)
    excy = excy + 1
    End If
Next testexcepsy
excepsy.num_excep = excy
excepsy.xdist(excy) = ldist
excepsy.elev(excy) = nexelevy.elev(0)

' find all the exceps on left face ---------------------------------------------
For testexcepsx = 0 To curelev.num_excep
    If curelev.ydist(testexcepsx) = 0 Then
    excepsx.xdist(excx) = curelev.xdist(testexcepsx)
    excepsx.elev(excx) = curelev.elev(testexcepsx)
    excx = excx + 1
    End If
Next testexcepsx
excepsx.num_excep = excx
excepsx.xdist(excx) = wdist
excepsx.elev(excx) = nexelevx.elev(0)

' find front and right edge
' find all the exceps on front face --------------------------------------------
For testexcepsny = 0 To nexelevx.num_excep
    If nexelevx.xdist(testexcepsny) = 0 Then
    excepsfront.xdist(excf) = nexelevx.ydist(testexcepsny) 'y
    excepsfront.elev(excf) = nexelevx.elev(testexcepsny)
    excf = excf + 1
    End If
Next testexcepsny
excepsfront.num_excep = excf
excepsfront.xdist(excf) = ldist
excepsfront.elev(excf) = nexelevxy.elev(0)

' find all the exceps on right face ---------------------------------------------
For testexcepsnx = 0 To nexelevy.num_excep
    If nexelevy.ydist(testexcepsnx) = 0 Then
    excepsright.xdist(excr) = nexelevy.xdist(testexcepsnx)
    excepsright.elev(excr) = nexelevy.elev(testexcepsnx)
    excr = excr + 1
    End If
Next testexcepsnx
excepsright.num_excep = excr
excepsright.xdist(excr) = wdist
excepsright.elev(excr) = nexelevxy.elev(0)

' line up other exceps average distance on each axis
' put 0 excep in here
' ******* error is here due xdist(n-1)=xdist(n)
' also investigate 0 elev in excepsx

' average dist of exceptions in each axis
' calc sub_elev at average dist -> set to .elev(0)
' find all non-axial exceps if any -> .elev(1 to n)
' calc sub_elev on next face at average dist -> .elev(n+1)

For testmdexcepsx = 0 To curelev.num_excep
    If curelev.xdist(testmdexcepsx) > 0 And curelev.ydist(testmdexcepsx) > 0 Then
    excmx = excmx + 1
'    Stop
    mdexcepsy.elev(excmx) = curelev.elev(testmdexcepsx)
    mdexcepsy.xdist(excmx) = curelev.xdist(testmdexcepsx)
    mdexcepsy.ydist(0) = mdexcepsy.ydist(0) + curelev.ydist(testmdexcepsx) ' y should be averaged
    End If
Next testmdexcepsx
If excmx > 0 Then mdflag = True
excmx = excmx + 1
mdexcepsy.num_excep = excmx
mdexcepsy.xdist(excmx) = ldist ' can't all be ldist
mdexcepsy.ydist(0) = mdexcepsy.ydist(0) / excmx
mdexcepsy.xdist(0) = 0
' 0 and last elevs should be sub_elev at y distance
' problem here which generates divide by zero
mdexcepsy.elev(0) = sub_elev(nexelevy.elev(0), nexelevxy.elev(0), mdexcepsy.ydist(0), ldist)
mdexcepsy.elev(excmx) = sub_elev(curelev.elev(0), nexelevy.elev(0), mdexcepsy.ydist(0), ldist)
' put 0 excep (origin sub_elev) here
If mdexcepsx.xdist(1) > 0 Then Stop
Debug.Print "++"; mdexcepsx.xdist(1), mdexcepsx.xdist(2)
For testmdexcepsy = 0 To curelev.num_excep
   If curelev.xdist(testmdexcepsy) > 0 And curelev.ydist(testmdexcepsy) > 0 Then
   excmy = excmy + 1
'Stop
    mdexcepsx.elev(excmy) = curelev.elev(testmdexcepsy)
     mdexcepsx.xdist(excmy) = curelev.ydist(testmdexcepsy) '
     Debug.Print "mxx"; mdexcepsx.xdist(excmy)
     mdexcepsx.ydist(0) = mdexcepsx.xdist(0) + curelev.xdist(testmdexcepsy) ' should be averaged
     End If
Next testmdexcepsy
excmy = excmy + 1
mdexcepsx.num_excep = excmy
Debug.Print "excmy"; excmy
mdexcepsx.xdist(excmy) = wdist
Debug.Print "mxxW"; mdexcepsx.xdist(excmy)
mdexcepsx.ydist(0) = mdexcepsx.ydist(0) / excmy
mdexcepsx.elev(0) = sub_elev(nexelevx.elev(0), nexelevxy.elev(0), mdexcepsx.xdist(0), wdist)
' 0 and last elevs should be sub_elev at x distance
mdexcepsx.elev(excmy) = sub_elev(curelev.elev(0), nexelevx.elev(0), mdexcepsx.xdist(0), wdist)
mdexcepsx.xdist(0) = 0
'If mdexcepsx.xdist(1) = mdexcepsx.xdist(2) Then Stop

If cutorfill Then GoTo cut Else GoTo fill
cut:
If mdflag Then
avex = (ce_area(excepsx, wdist) + ce_area(mdexcepsx, wdist) + ce_area(excepsright, wdist)) / 3
avey = (ce_area(excepsy, ldist) + ce_area(mdexcepsy, ldist) + ce_area(excepsfront, ldist)) / 3
Else
avex = (ce_area(excepsx, wdist) + ce_area(excepsright, wdist)) / 2
avey = (ce_area(excepsy, ldist) + ce_area(excepsfront, ldist)) / 2
End If
volx = avex * ldist
voly = avey * wdist
avol = (volx + voly) / 2
'Printer.Print "volx: "; volx; " voly: "; voly; " avol/27: "; avol / 27
calc_cube = avol / 27
Exit Function
fill:
If mdflag Then
avex = (fe_area(excepsx, wdist) + fe_area(mdexcepsx, wdist) + fe_area(excepsright, wdist)) / 3
avey = (fe_area(excepsy, ldist) + fe_area(mdexcepsy, ldist) + fe_area(excepsfront, ldist)) / 3
Else
avex = (fe_area(excepsx, wdist) + fe_area(excepsright, wdist)) / 2
avey = (fe_area(excepsy, ldist) + fe_area(excepsfront, ldist)) / 2
End If
volx = avex * ldist
Debug.Print "volx: "; volx
voly = avey * wdist
Debug.Print "voly: "; voly
avol = (volx + voly) / 2
Debug.Print "avol: "; avol
'Printer.Print "volx: "; volx; " voly: "; voly; " avol/27: "; avol / 27
calc_cube = avol / 27
End Function
Public Function ce_area(face As Elevation, dist As Single) As Single
Debug.Print "ce area"
'Printer.Print "ce_area"
' calculate end area from pre-arranged elevation collection and station distance
' sort the elevations as distances
Dim distorx As Single, distory As Single, elevor As Single
'Stop
gap = Int(face.num_excep / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For Index = 0 To face.num_excep - gap
    If face.xdist(Index) = face.xdist(Index + gap) Then Stop ' test elsewhere
    If face.xdist(Index) > face.xdist(Index + gap) Then
     elevor = face.elev(Index)
     face.elev(Index) = face.elev(Index + gap)
     face.elev(Index + gap) = elevor
     distorx = face.xdist(Index)
     face.xdist(Index) = face.xdist(Index + gap)
     face.xdist(Index + gap) = distorx
     distory = face.ydist(Index)
     face.ydist(Index) = face.ydist(Index + gap)
     face.ydist(Index + gap) = distory
      doneflag = 0
     End If
    Next Index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
Dim subne As Single, sub_td As Single
Debug.Print "num exceps:"; face.num_excep
'Printer.Print "num_exceps: "; face.num_excep
For ne = 1 To face.num_excep
 sub_td = (face.xdist(ne) - face.xdist(ne - 1))
 'Printer.Print "sub_td: "; sub_td

  For subloop = 1 To 400

   curel = standard - sub_elev(face.elev(ne - 1), face.elev(ne), subne, sub_td)
   If curel > 0 Then totelev = totelev + curel: subne = subne + 0.0025 * sub_td
  'Debug.Print " s:"; subne; " c:"; curel; " t:"; totelev;
  'Printer.Print " s"; subne; " c"; curel; " t"; totelev;
  Next subloop
  Debug.Print
  'Printer.Print
   totdist = totdist + subne
   'Printer.Print "totdist: "; totdist
Next ne
Debug.Print "area: "; (totdist * totelev) * shrink
'Printer.Print "area: "; (totdist * totelev) * shrink
ce_area = (totdist * totelev) * shrink
End Function
Public Function fe_area(face As Elevation, dist As Single) As Single
Debug.Print "fe area"
'Printer.Print "fe_area"
Dim distorx As Single, distory As Single, elevor As Single
Dim totelev As Single, totdist As Single
gap = Int(face.num_excep / 2)
  Do While gap >= 1
   Do
   doneflag = 1
    For Index = 0 To face.num_excep - gap
    ' test for this elsewhere/in inputs? runtime failure when compiled
     If face.xdist(Index) = face.xdist(Index + gap) Then Stop
     If face.xdist(Index) > face.xdist(Index + gap) Then
     elevor = face.elev(Index)
     face.elev(Index) = face.elev(Index + gap)
     face.elev(Index + gap) = elevor
      distorx = face.xdist(Index)
     face.xdist(Index) = face.xdist(Index + gap)
     face.xdist(Index + gap) = distorx
     distory = face.ydist(Index)
     face.ydist(Index) = face.ydist(Index + gap)
     face.ydist(Index + gap) = distory
     doneflag = 0
     End If
    Next Index
   Loop Until doneflag = 1
   gap = Int(gap / 2)
  Loop
Dim subne As Single, sub_td As Single
Debug.Print face.num_excep
'Printer.Print "numexceps: "; face.num_excep
For ne = 1 To face.num_excep
  sub_td = (face.xdist(ne) - face.xdist(ne - 1)) ' tenths error could be here if xdist wrong
  Debug.Print "sub_td:"; sub_td
  'Printer.Print "sub_td: "; sub_td
  For subloop = 1 To 400
  ' replace standard with subelev of benchplane
  curel = standard - sub_elev(face.elev(ne - 1), face.elev(ne), subne, sub_td)
  If curel < 0 Then totelev = totelev + Abs(curel): subne = subne + 0.0025 * sub_td
  'Debug.Print " s:"; subne; " c:"; curel; " t:"; totelev;
  'Printer.Print " s"; subne; " c"; curel; " t"; totelev;
  Next subloop
  'Printer.Print
  totdist = totdist + subne
Next ne
fe_area = (totdist * totelev) * shrink
End Function
Public Function sub_elev(ep1 As Single, ep2 As Single, subdist As Single, totaldist As Single) As Single
If totaldist = 0 Then Stop: Beep 'Exit Function
Dim slope As Single
    difhyt = -1 * (ep1 - ep2)
    slope = difhyt / totaldist
sub_elev = ep1 + subdist * slope
End Function
