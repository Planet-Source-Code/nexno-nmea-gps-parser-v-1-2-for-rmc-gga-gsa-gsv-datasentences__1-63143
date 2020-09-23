Attribute VB_Name = "NMEA"
'by NexNo
'contact nexno@inetmx.de
'website http://www.qudi.de

Option Explicit
Option Base 1

'The RMC-Datasentence (RMC=recommended minimum sentence C)
'is a recommendation for the minimum, that a GPS-Receiver should give back.
'It looks like this: "$GPRMC,191410,A,4735.5634,N,00739.3538,E,0.0,0.0,181102,0.4,E,A*19"
Public Sub decodeRMC(ByVal inp As String, _
            Optional ByRef UtcTime As String, _
            Optional ByRef ReceiverWarning As Boolean, _
            Optional ByRef Latitude As Double, _
            Optional ByRef LatitudeDir As String, _
            Optional ByRef Longitude As Double, _
            Optional ByRef LongitudeDir As String, _
            Optional ByRef SpeedKMH As Double, _
            Optional ByRef Course As Double, _
            Optional ByRef DateStamp As String, _
            Optional ByRef MagneticDeclination As Double, _
            Optional ByRef Checksum As Boolean)
    On Error Resume Next
    inp = UCase(Trim(inp))
    'Checking initstring. Must be the same for all RMC sentences.
    If Left(inp, 1) <> "$" Or Mid(inp, 4, 3) <> "RMC" Then Checksum = False: Exit Sub
    'Extracting that part of the sentence that is needed to calculate the checksum
    Dim ChkDat As String
    ChkDat = Mid(inp, 2, InStr(2, inp, "*") - 2)
    'For compatibility with split function
    inp = Replace(inp, ",,", ", ,")
    'Splitting sentence
    Dim Dat As Variant
    Dat = Split(inp, ",")
    'Calculating checksum and comparing it
    Dim ChkSum As String
    ChkSum = Dat(UBound(Dat))
    ChkSum = Right(ChkSum, Len(ChkSum) - InStr(1, ChkSum, "*"))
    If calcChecksum(ChkDat) = Hex2Dec(ChkSum) Then Checksum = True Else Checksum = False: Exit Sub
    'UtcTime
    UtcTime = Left(Dat(2), 6)
    If UtcTime <> " " Then UtcTime = Left(UtcTime, 2) & ":" & Mid(UtcTime, 3, 2) & ":" & Right(UtcTime, 2) Else UtcTime = ""
    'ReceiverWarning
    If Dat(3) = "A" Or Dat(3) = "" Then ReceiverWarning = False Else ReceiverWarning = True
    'Length
    Dim sp As Integer
    sp = InStr(1, Dat(4), ".")
    Latitude = CDbl(Left(Dat(4), sp - 3)) + CDbl(CDbl(Replace(Mid(Dat(4), sp - 2), ".", ",")) / 60)
    Latitude = Round(Latitude, 8)
    'LengthDir
    LatitudeDir = Dat(5)
    'Width
    sp = InStr(1, Dat(6), ".")
    Longitude = CDbl(Left(Dat(6), sp - 3)) + CDbl(CDbl(Replace(Mid(Dat(6), sp - 2), ".", ",")) / 60)
    Longitude = Round(Longitude, 8)
    'WidthDir
    LongitudeDir = Dat(7)
    'SpeedKMH (needs to be converted from knots)
    SpeedKMH = Replace(CStr((Dat(8) * 0.54)), ".", ",")
    'Course without movement
    Course = Replace(CStr(Dat(9)), ".", ",")
    'DateStamp
    DateStamp = Left(Dat(10), 6)
    If DateStamp <> " " Then DateStamp = Left(DateStamp, 2) & "." & Mid(DateStamp, 3, 2) & "." & Mid(DateStamp, 5) Else DateStamp = ""
    'MagneticDeclination
    MagneticDeclination = Replace(CStr(Dat(11)), ".", ",")
End Sub

'The GGA-Datasentence contains the most important information about GPS-position and accuracy.
'it looks like: "$GPGGA,191410,4735.5634,N,00739.3538,E,1,04,4.4,351.5,M,48.0,M,,*45"
Public Sub decodeGGA(ByVal inp As String, _
            Optional ByRef UtcTime As String, _
            Optional ByRef Latitude As Double, _
            Optional ByRef LatitudeDir As String, _
            Optional ByRef Longitude As Double, _
            Optional ByRef LongitudeDir As String, _
            Optional ByRef Quality As String, _
            Optional ByRef SatellitesIV As Integer, _
            Optional ByRef HDOP As Double, _
            Optional ByRef AltitudeSea As Double, _
            Optional ByRef AltitudeSeaUnit As String, _
            Optional ByRef AltitudeEllipsoid As Double, _
            Optional ByRef AltitudeEllipsoidUnit As String, _
            Optional ByRef Checksum As Boolean)
    On Error Resume Next
    inp = UCase(Trim(inp))
    'Checking initstring. Must be the same for all GGA sentences.
    If Left(inp, 1) <> "$" Or Mid(inp, 4, 3) <> "GGA" Then Checksum = False: Exit Sub
    'Extracting that part of the sentence that is needed to calculate the checksum
    Dim ChkDat As String
    ChkDat = Mid(inp, 2, InStr(2, inp, "*") - 2)
    'For compatibility with split function
    inp = Replace(inp, ",,", ", ,")
    'Splitting sentence
    Dim Dat As Variant
    Dat = Split(inp, ",")
    'Calculating checksum and comparing it
    Dim ChkSum As String
    ChkSum = Dat(UBound(Dat))
    ChkSum = Right(ChkSum, Len(ChkSum) - InStr(1, ChkSum, "*"))
    If calcChecksum(ChkDat) = Hex2Dec(ChkSum) Then Checksum = True Else Checksum = False: Exit Sub
    'UtcTime
    UtcTime = Left(Dat(2), 6)
    If UtcTime <> " " Then UtcTime = Left(UtcTime, 2) & ":" & Mid(UtcTime, 3, 2) & ":" & Right(UtcTime, 2) Else UtcTime = ""
    'Length
    Dim sp As Integer
    sp = InStr(1, Dat(3), ".")
    Latitude = CDbl(Left(Dat(3), sp - 3)) + CDbl(CDbl(Replace(Mid(Dat(3), sp - 2), ".", ",")) / 60)
    Latitude = Round(Latitude, 8)
    'LengthDir
    LatitudeDir = Dat(4)
    'Width
    sp = InStr(1, Dat(5), ".")
    Longitude = CDbl(Left(Dat(5), sp - 3)) + CDbl(CDbl(Replace(Mid(Dat(5), sp - 2), ".", ",")) / 60)
    Longitude = Round(Longitude, 8)
    'WidthDir
    LongitudeDir = Dat(6)
    'Quality: 0-invalid, 1-gps, 2-dgps, 6-guessed
    Quality = "unknown"
    If Dat(7) = 0 Then Quality = "no fix"
    If Dat(7) = 1 Then Quality = "GPS fix"
    If Dat(7) = 2 Then Quality = "DGPS fix"
    If Dat(7) = 6 Then Quality = "guessed"
    'Satellites in view
    SatellitesIV = Dat(8)
    'HDOP: horizontal dilution of precision (accuracy)
    HDOP = Replace(CStr(Dat(9)), ".", ",")
    'Altitude over sea
    AltitudeSea = Replace(CStr(Dat(10)), ".", ",")
    'Altitude over sea unit
    AltitudeSeaUnit = Dat(11)
    'Altitude over ellipsoid
    AltitudeEllipsoid = Replace(CStr(Dat(12)), ".", ",")
    'Altitude over ellipsoid unit
    AltitudeEllipsoidUnit = Dat(13)
End Sub

'The GSA-Datasentence contains information about the PRN-Numbers of the satellites that are used
'for calculating the actual position and some more detailed info about the accuracy.
'it looks like: "$GPGSA,A,3,,,,15,17,18,23,,,,,,4.7,4.4,1.5*3F"
Public Sub decodeGSA(ByVal inp As String, _
            Optional ByRef AutoSel As Boolean, _
            Optional ByRef mode As String, _
            Optional ByRef prn As Variant, _
            Optional ByRef PDOP As Double, _
            Optional ByRef HDOP As Double, _
            Optional ByRef VDOP As Double, _
            Optional ByRef Checksum As Boolean)
    On Error Resume Next
    inp = UCase(Trim(inp))
    'Checking initstring. Must be the same for all GSA sentences.
    If Left(inp, 1) <> "$" Or Mid(inp, 4, 3) <> "GSA" Then Checksum = False: Exit Sub
    'Extracting that part of the sentence that is needed to calculate the checksum
    Dim ChkDat As String
    ChkDat = Mid(inp, 2, InStr(2, inp, "*") - 2)
    'For compatibility with split function
    inp = Replace(inp, ",,", ", ,")
    'Splitting sentence
    Dim Dat As Variant
    Dat = Split(inp, ",")
    'Calculating checksum and comparing it
    Dim ChkSum As String
    ChkSum = Dat(UBound(Dat))
    ChkSum = Right(ChkSum, Len(ChkSum) - InStr(1, ChkSum, "*"))
    If calcChecksum(ChkDat) = Hex2Dec(ChkSum) Then Checksum = True Else Checksum = False: Exit Sub
    'Auto Selection Mode
    If Dat(2) = "A" Then AutoSel = True Else AutoSel = False
    'Mode
    mode = "unknown"
    If Dat(3) = "3" Then mode = "3D-Fix"
    If Dat(3) = "2" Then mode = "2D-Fix"
    If Dat(3) = "1" Then mode = "No-Fix"
    'PRN-Numbers
    ReDim prn(12)
    Dim i As Integer
    For i = 4 To 15
        If IsNumeric(Dat(i)) Then prn(i - 3) = CInt(Dat(i)) Else prn(i - 3) = ""
    Next
    'PDOP in meters
    PDOP = Replace(CStr(Dat(16)), ".", ",")
    'HDOP horizontal dilution of precision in meters
    HDOP = Replace(CStr(Dat(17)), ".", ",")
    'VDOP vertical dilution of precision in meters
    VDOP = Replace(Left(CStr(Dat(18)), InStr(1, Dat(18), "*") - 1), ".", ",")
End Sub

'              sn ed azi dB ------------ ------------ ------------
'$GPGSV,3,1,12,22,89,000,00,14,59,000,00,15,53,000,00,18,51,000,00*7F
'              ------------ ------------ ---------- ----------
'$GPGSV,3,2,12,09,47,000,00,19,15,000,00,21,13,000,,31,13,000,*7C
'              ------------ ---------- ------------ ------------
'$GPGSV,3,3,12,03,06,000,00,11,04,000,,28,02,000,00,05,01,000,00*77

'$GPGSV,3,1,12,22,89,000,00,14,59,000,00,15,53,000,00,18,51,000,00*7F
'$GPGSV,3,2,12,09,47,000,00,19,15,000,00,21,13,000,,31,13,000,*7C
'$GPGSV,3,3,12,03,06,000,00,11,04,000,,28,02,000,00,05,01,000,00*77
'GSV - Satellites in view
'        1 2 3 4 5 6 7     n
'        | | | | | | |     |
' $--GSV,x,x,x,x,x,x,x,...*hh<CR><LF>
'  1) total number of messages
'  2) message number
'  3) satellites in view
'  4) satellite number
'  5) elevation in degrees
'  6) azimuth in degrees to true
'  7) SNR in dB
'  more satellite infos like 4)-7)
'  n) checksum
  
Public Sub decodeGSV(ByVal inp As String, _
            Optional ByRef tnm As Integer, _
            Optional ByRef mn As Integer, _
            Optional ByRef SatsInView As Integer, _
            Optional ByRef SatNr As Variant, _
            Optional ByRef Elevation As Variant, _
            Optional ByRef Azimuth As Variant, _
            Optional ByRef SNRdB As Variant, _
            Optional ByRef Checksum As Boolean)
    On Error Resume Next
    inp = UCase(Trim(inp))
    'Checking initstring. Must be the same for all GSV sentences.
    If Left(inp, 1) <> "$" Or Mid(inp, 4, 3) <> "GSV" Then Checksum = False: Exit Sub
    'Extracting that part of the sentence that is needed to calculate the checksum
    Dim ChkDat As String
    ChkDat = Mid(inp, 2, InStr(2, inp, "*") - 2)
    ChkDat = ChkDat
    'For compatibility with split function
    inp = Replace(inp, ",,", ", ,")
    'Splitting sentence
    Dim Dat As Variant
    Dat = Split(inp, ",")
    'Cutting what we won't need
    inp = Right(inp, Len(inp) - 7)
    inp = Left(inp, InStr(1, inp, "*") - 1)
    'Calculating checksum and comparing it
    Dim ChkSum As String
    ChkSum = Dat(UBound(Dat))
    ChkSum = Right(ChkSum, Len(ChkSum) - InStr(1, ChkSum, "*"))
    Dim ccs As Integer, dcs As Long
    ccs = calcChecksum(ChkDat)
    dcs = Hex2Dec(ChkSum)
    If ccs = dcs Then Checksum = True Else Checksum = False: Exit Sub
    'total number of messages
    tnm = CInt(Dat(1))
    'message number
    mn = CInt(Dat(2))
    'Satellites in View
    SatsInView = Dat(3)
    Dim i As Integer
    Dim sp As Integer
    'Satellite-Numbers
    ReDim SatNr(4)
    For i = 1 To 4
        sp = InStr(1, Dat(3 + i), "*")
        If sp > 0 Then Dat(3 + i) = Left(Dat(3 + i), sp - 1)
        SatNr(i) = Dat(3 + i)
    Next
    'Elevations
    ReDim Elevation(4)
    For i = 1 To 4
        sp = InStr(1, Dat(7 + i), "*")
        If sp > 0 Then Dat(7 + i) = Left(Dat(7 + i), sp - 1)
        Elevation(i) = Dat(7 + i)
    Next
    'Azimuth's
    ReDim Azimuth(4)
    For i = 1 To 4
        sp = InStr(1, Dat(11 + i), "*")
        If sp > 0 Then Dat(11 + i) = Left(Dat(11 + i), sp - 1)
        Azimuth(i) = Dat(11 + i)
    Next
    'SNRdB's
    ReDim SNRdB(4)
    For i = 1 To 4
        sp = InStr(1, Dat(15 + i), "*")
        If sp > 0 Then Dat(15 + i) = Left(Dat(15 + i), sp - 1)
        SNRdB(i) = Dat(15 + i)
    Next
End Sub
  
  
  
'Helper functions:
'=================

Function calcChecksum(inp As String) As Integer
    Dim i As Integer, s As Integer
    s = 0
    For i = 1 To Len(inp)
        s = s Xor Asc(Mid(inp, i, 1))
    Next
    calcChecksum = s
End Function

Function Hex2Dec(HexNum As Variant) As Long
    Hex2Dec = "&h" & HexNum
End Function


'Compatibility functions:
'========================

Function Split(sIn As String, sDel As String) As Variant
    Dim i As Integer, x As Integer, s As Integer, t As Integer
    i = 1: s = 1: t = 1: x = 1
    ReDim tArr(1 To x) As Variant
    If InStr(1, sIn, sDel) <> 0 Then
        Do
            ReDim Preserve tArr(1 To x) As Variant
            tArr(i) = Mid(sIn, t, InStr(s, sIn, sDel) - t)
            t = InStr(s, sIn, sDel) + Len(sDel)
            s = t
            If tArr(i) <> "" Then i = i + 1
            x = x + 1
        Loop Until InStr(s, sIn, sDel) = 0
        ReDim Preserve tArr(1 To x) As Variant
        tArr(i) = Mid(sIn, t, Len(sIn) - t + 1)
    Else
        tArr(1) = sIn
    End If
    Split = tArr
End Function

Function Round(ByVal Value As Variant, Optional ByVal digits As Integer = 0) As Variant
  Dim i As Long
  Dim Pot10(-28 To 28) As Variant
  If i = 0 Then
    For i = LBound(Pot10) To UBound(Pot10)
      Pot10(i) = CDec(10 ^ i)
    Next i
  End If
  On Error Resume Next
    If Value > 0 Then
      Round = Int(Value * Pot10(digits) + 0.5) * Pot10(-digits)
    Else
      Round = -Int(-Value * Pot10(digits) + 0.5) * Pot10(-digits)
    End If
    If Err.Number Then Round = Value
  On Error GoTo 0
End Function

Function Replace(strString As String, Find As String, strReplace As String) As String
    Dim ss As Long
    ss = InStr(1, strString, Find)
    If ss > 0 Then
        strString = Left(strString, ss - 1) & strReplace & Right(strString, Len(strString) - (ss + (Len(Find) - 1)))
        Replace = Replace(strString, Find, strReplace)
    Else
        Replace = strString
    End If
End Function

