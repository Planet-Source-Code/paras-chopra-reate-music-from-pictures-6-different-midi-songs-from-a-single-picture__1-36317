VERSION 5.00
Object = "{B6C1EA38-375B-11D4-93AB-E7C32384627A}#3.0#0"; "FREELIB.OCX"
Begin VB.Form Form1 
   Caption         =   "Create music from pictures"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Play song number 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Play song number 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play song number 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play song number 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play song number 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play song number 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   1560
      Width           =   540
   End
   Begin FreeLibSrc.FreeLib FreeLib1 
      Height          =   480
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      IniFile         =   ""
      IniSection      =   ""
      IniSize         =   ""
      IniDefault      =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   75
   End
   Begin VB.Menu a 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Sub a_Click()
o = MsgBox("Created by Paras Chopra, paraschopra@lycos.com, http://naramcheez.netfirms.com. Do you want to see the readme?", vbYesNo)
If o = vbYes Then
Shell "start " & App.Path & "\readme.txt"
End If
End Sub

Private Sub Command1_Click()
play1
End Sub

Sub play1()
FreeLib1.MidiChannel = 0
FreeLib1.MidiPan = 63
FreeLib1.MidiPortOpen = True
For X = 0 To Picture1.ScaleWidth
DoEvents
For Y = 0 To Picture1.ScaleHeight
DoEvents
buffer = GetPixel(Picture1.hdc, X, Y)
blue = Int(buffer / 65536)
green = Int((buffer - blue * 65536) / 256)
red = Int((buffer - blue * 65536) - (green * 256))
redlab:
If red > 127 Then
red = red - 127
GoTo redlab
End If
greenlab:
If green > 127 Then
green = green - 127
GoTo greenlab
End If
bluelab:
If red > 127 Then
blue = blue - 127
GoTo bluelab
End If
FreeLib1.MidiInstrument = red
FreeLib1.MidiVolume = green
FreeLib1.MidiNotePlay = blue
Sleep (100)
Next Y
Next X
FreeLib1.MidiPortOpen = False
End Sub

Private Sub Command2_Click()
FreeLib1.MidiChannel = 0
FreeLib1.MidiPan = 63
FreeLib1.MidiPortOpen = True
For X = 0 To Picture1.ScaleWidth
DoEvents
For Y = 0 To Picture1.ScaleHeight
DoEvents
buffer = GetPixel(Picture1.hdc, X, Y)
blue = Int(buffer / 65536)
green = Int((buffer - blue * 65536) / 256)
red = Int((buffer - blue * 65536) - (green * 256))
redlab:
If red > 127 Then
red = red - 127
GoTo redlab
End If
greenlab:
If green > 127 Then
green = green - 127
GoTo greenlab
End If
bluelab:
If red > 127 Then
blue = blue - 127
GoTo bluelab
End If
FreeLib1.MidiInstrument = red
FreeLib1.MidiVolume = blue
FreeLib1.MidiNotePlay = green
Sleep (100)
Next Y
Next X
FreeLib1.MidiPortOpen = False
End Sub

Private Sub Command3_Click()
FreeLib1.MidiChannel = 0
FreeLib1.MidiPan = 63
FreeLib1.MidiPortOpen = True
For X = 0 To Picture1.ScaleWidth
DoEvents
For Y = 0 To Picture1.ScaleHeight
DoEvents
buffer = GetPixel(Picture1.hdc, X, Y)
blue = Int(buffer / 65536)
green = Int((buffer - blue * 65536) / 256)
red = Int((buffer - blue * 65536) - (green * 256))
redlab:
If red > 127 Then
red = red - 127
GoTo redlab
End If
greenlab:
If green > 127 Then
green = green - 127
GoTo greenlab
End If
bluelab:
If red > 127 Then
blue = blue - 127
GoTo bluelab
End If
FreeLib1.MidiInstrument = green
FreeLib1.MidiVolume = red
FreeLib1.MidiNotePlay = blue
Sleep (100)
Next Y
Next X
FreeLib1.MidiPortOpen = False
End Sub

Private Sub Command4_Click()
FreeLib1.MidiChannel = 0
FreeLib1.MidiPan = 63
FreeLib1.MidiPortOpen = True
For X = 0 To Picture1.ScaleWidth
DoEvents
For Y = 0 To Picture1.ScaleHeight
DoEvents
buffer = GetPixel(Picture1.hdc, X, Y)
blue = Int(buffer / 65536)
green = Int((buffer - blue * 65536) / 256)
red = Int((buffer - blue * 65536) - (green * 256))
redlab:
If red > 127 Then
red = red - 127
GoTo redlab
End If
greenlab:
If green > 127 Then
green = green - 127
GoTo greenlab
End If
bluelab:
If red > 127 Then
blue = blue - 127
GoTo bluelab
End If
FreeLib1.MidiInstrument = green
FreeLib1.MidiVolume = blue
FreeLib1.MidiNotePlay = red
Sleep (100)
Next Y
Next X
FreeLib1.MidiPortOpen = False
End Sub

Private Sub Command5_Click()
FreeLib1.MidiChannel = 0
FreeLib1.MidiPan = 63
FreeLib1.MidiPortOpen = True
For X = 0 To Picture1.ScaleWidth
DoEvents
For Y = 0 To Picture1.ScaleHeight
DoEvents
buffer = GetPixel(Picture1.hdc, X, Y)
blue = Int(buffer / 65536)
green = Int((buffer - blue * 65536) / 256)
red = Int((buffer - blue * 65536) - (green * 256))
redlab:
If red > 127 Then
red = red - 127
GoTo redlab
End If
greenlab:
If green > 127 Then
green = green - 127
GoTo greenlab
End If
bluelab:
If red > 127 Then
blue = blue - 127
GoTo bluelab
End If
FreeLib1.MidiInstrument = blue
FreeLib1.MidiVolume = red
FreeLib1.MidiNotePlay = green
Sleep (100)
Next Y
Next X
FreeLib1.MidiPortOpen = False
End Sub

Private Sub Command6_Click()
FreeLib1.MidiChannel = 0
FreeLib1.MidiPan = 63
FreeLib1.MidiPortOpen = True
For X = 0 To Picture1.ScaleWidth
DoEvents
For Y = 0 To Picture1.ScaleHeight
DoEvents
buffer = GetPixel(Picture1.hdc, X, Y)
blue = Int(buffer / 65536)
green = Int((buffer - blue * 65536) / 256)
red = Int((buffer - blue * 65536) - (green * 256))
redlab:
If red > 127 Then
red = red - 127
GoTo redlab
End If
greenlab:
If green > 127 Then
green = green - 127
GoTo greenlab
End If
bluelab:
If red > 127 Then
blue = blue - 127
GoTo bluelab
End If
FreeLib1.MidiInstrument = blue
FreeLib1.MidiVolume = green
FreeLib1.MidiNotePlay = red
Sleep (100)
Next Y
Next X
FreeLib1.MidiPortOpen = False
End Sub

Private Sub Command7_Click()
FreeLib1.MidiPortOpen = False
End Sub

Private Sub Form_Load()
Label1.Caption = "The length of the song is " & (Picture1.ScaleHeight * Picture1.ScaleWidth) / 10 & " seconds."
End Sub
