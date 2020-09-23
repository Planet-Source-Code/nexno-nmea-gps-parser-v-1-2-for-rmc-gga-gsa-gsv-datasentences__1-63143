VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "NMEA-Parser v.1.2"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame3 
      Caption         =   "GSA-Data"
      Height          =   3255
      Left            =   6240
      TabIndex        =   44
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   12
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   11
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   10
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   9
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   8
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   7
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   5
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   4
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   2
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPRN 
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtVDOP 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtHDOP 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtPDOP 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtMO 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtASM 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "PRN numbers of satellites:"
         Height          =   285
         Left            =   120
         TabIndex        =   55
         ToolTipText     =   "PRN numbers of satellites currently used"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "VDOP:"
         Height          =   285
         Left            =   120
         TabIndex        =   53
         ToolTipText     =   "Vertical delution of precision"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "HDOP:"
         Height          =   285
         Left            =   120
         TabIndex        =   51
         ToolTipText     =   "Horisontal dilution of precision"
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "PDOP:"
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Mode:"
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "AutoSel:"
         Height          =   285
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "Auto Selection Mode"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "GGA-Data"
      Height          =   3255
      Left            =   3120
      TabIndex        =   23
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtALEU 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox txtALSU 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtALE 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtALS 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtHD 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtSA 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtQU 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtLOD2 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtLAD2 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtUT2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtLA2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtLO2 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "AltEllipsoid:"
         Height          =   285
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "Altitude over Ellipsoid"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "AltSea:"
         Height          =   285
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Altitude over Sea"
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "HDOP:"
         Height          =   285
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "horizontal dilution of precision"
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "SatellitesIV:"
         Height          =   285
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Sat. in view"
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Quality:"
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Longitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Latitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "UtcTime:"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RMC-Data"
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox txtMD 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtDS 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtCO 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtSK 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtLAD 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtLOD 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtLO 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtLA 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtRW 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtUT 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "MagDec:"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Magnetic Declination"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Date:"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Course:"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Speed:"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Latitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Longitude:"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "RecWarn:"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "UtcTime:"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&start"
      Default         =   -1  'True
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtComPort 
      Height          =   285
      Left            =   8880
      TabIndex        =   2
      Text            =   "5"
      Top             =   0
      Width           =   375
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8640
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Frame Frame4 
      Caption         =   "GSV-Data"
      Height          =   1815
      Left            =   0
      TabIndex        =   68
      Top             =   4920
      Width           =   9255
      Begin VB.Label Label24 
         Alignment       =   2  'Zentriert
         Caption         =   "GSV-parser is included in the module, but I didn't had time for the gui-connection right now."
         Height          =   255
         Left            =   600
         TabIndex        =   69
         Top             =   840
         Width           =   7815
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "ComPort:"
      Height          =   285
      Left            =   7920
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buffer As String, x As String, _
    rmc As String, ut As String, rw As Boolean, la As Double, lad As String, _
    lo As Double, lod As String, sk As Double, co As Double, ds As String, _
    md As Double, cs As Boolean, _
    gga As String, qu As String, sa As Integer, hd As Double, als As Double, _
    alsu As String, ale As Double, aleu As String, _
    gsa As String, asm As Boolean, mo As String, prn As Variant, pd As Double, vd As Double, _
    gsv As String, tnm As Integer, mn As Integer, siv As Integer, satn As Variant, ele As Variant, azi As Variant, snr As Variant, _
    gsvs As Variant

            
Private Sub Command1_Click()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        Command1.Caption = "&start"
    Else
        MSComm1.CommPort = txtComPort.Text
        MSComm1.Settings = "4800,n,8,1"
        MSComm1.RThreshold = 1
        MSComm1.PortOpen = True
        Command1.Caption = "&stop"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > 6000 Then Text1.Text = Right(Text1.Text, 3000)
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub MSComm1_OnComm()
If MSComm1.CommEvent = comEvReceive Then
    x = MSComm1.Input
    Buffer = Buffer & x
    'If InStr(1, Buffer, vbCrLf) Then Text2.Text = Text2.Text & Buffer
    'Me.Caption = Len(Buffer)
    
    'Is there a completed GPRMC Datasentence in the buffer?
    rmcS = InStr(1, Buffer, "$GPRMC")
    If rmcS > 0 Then
        rmcE = InStr(rmcS, Buffer, vbCrLf)
        If rmcE > 0 Then 'GPRMC Datasentence found
            rmc = Mid(Buffer, rmcS, rmcE - rmcS)
            decodeRMC rmc, ut, rw, la, lad, lo, lod, sk, co, ds, md, cs
            If cs Then 'checksum is correct
                Text1.Text = Text1.Text & rmc & vbCrLf
                txtUT.Text = ut
                txtRW.Text = rw
                txtLA.Text = la
                txtLAD.Text = lad
                txtLO.Text = lo
                txtLOD.Text = lod
                txtSK.Text = sk
                txtCO.Text = co
                txtDS.Text = ds
                txtMD.Text = md
            End If
            Buffer = Right(Buffer, Len(Buffer) - rmcE) 'remove parsed data from the buffer
        End If
    End If
    
    'Is there a completed GPGGA Datasentence in the buffer?
    ggaS = InStr(1, Buffer, "$GPGGA")
    If ggaS > 0 Then
        ggaE = InStr(ggaS, Buffer, vbCrLf)
        If ggaE > 0 Then 'GPGGA Datasentence found
            gga = Mid(Buffer, ggaS, ggaE - ggaS)
            decodeGGA gga, ut, la, lad, lo, lod, qu, sa, hd, als, alsu, ale, aleu, cs
            If cs Then 'checksum is correct
                Text1.Text = Text1.Text & gga & vbCrLf
                txtUT2.Text = ut
                txtLA2.Text = la
                txtLAD2.Text = lad
                txtLO2.Text = lo
                txtLOD2.Text = lod
                txtQU.Text = qu
                txtSA.Text = sa
                txtHD.Text = hd
                txtALS.Text = als
                txtALSU.Text = alsu
                txtALE.Text = ale
                txtALEU.Text = aleu
            End If
            Buffer = Right(Buffer, Len(Buffer) - ggaE) 'remove parsed data from the buffer
        End If
    End If
    
    'Is there a completed GPGSA Datasentence in the buffer?
    gsaS = InStr(1, Buffer, "$GPGSA")
    If gsaS > 0 Then
        gsaE = InStr(gsaS, Buffer, vbCrLf)
        If gsaE > 0 Then 'GPGSA Datasentence found
            gsa = Mid(Buffer, gsaS, gsaE - gsaS)
            decodeGSA gsa, asm, mo, prn, pd, hd, vd, cs
            If cs Then 'checksum is correct
                Text1.Text = Text1.Text & gsa & vbCrLf
                txtASM.Text = asm
                txtMO.Text = mo
                For i = 1 To UBound(prn)
                    txtPRN(i).Text = prn(i)
                Next
                txtPDOP.Text = pd
                txtHDOP.Text = hd
                txtVDOP.Text = vd
            End If
            Buffer = Right(Buffer, Len(Buffer) - gsaE) 'remove parsed data from the buffer
        End If
    End If
    
    'Is there a completed GPGSV Datasentence in the buffer?
    gsvs = InStr(1, Buffer, "$GPGSV")
    If gsvs > 0 Then
        gsvE = InStr(gsvs, Buffer, vbCrLf)
        If gsvE > 0 Then 'GPGSA Datasentence found
            gsv = Mid(Buffer, gsvs, gsvE - gsvs)
            Text1.Text = Text1.Text & gsv & vbCrLf
            decodeGSV gsv, tnm, mn, siv, satn, ele, azi, snr, cs
            If cs Then 'checksum is correct
                Text1.Text = Text1.Text & gsv & vbCrLf
                For i = 1 To 4
                    Text1.Text = Text1.Text & "[" & tnm & "] [" & mn & "] [" & siv & "] [" & satn(i) & "] [" & ele(i) & "] [" & azi(i) & "] [" & snr(i) & "]" & vbCrLf
                Next
            Else
                Text1.Text = Text1.Text & vbTab & "GSV-CHECKSUM INVALID" & vbCrLf
            End If
            Buffer = Right(Buffer, Len(Buffer) - gsvE) 'remove parsed data from the buffer
        End If
    End If
End If
End Sub
