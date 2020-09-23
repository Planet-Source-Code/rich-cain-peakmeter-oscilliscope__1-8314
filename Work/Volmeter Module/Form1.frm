VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3516
   ClientLeft      =   4236
   ClientTop       =   1620
   ClientWidth     =   4044
   LinkTopic       =   "Form1"
   ScaleHeight     =   3516
   ScaleWidth      =   4044
   Begin VB.PictureBox scopeBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFC0&
      Height          =   1812
      Left            =   120
      ScaleHeight     =   1764
      ScaleWidth      =   1404
      TabIndex        =   3
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2520
      Top             =   360
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim volume As Integer
Dim buffaddress As Long
Dim audbytearray As AUDINPUTARRAY
Dim retVal As Integer

Private Sub Form_Load()
    ProgressBar1.Max = 128
    ProgressBar1.Min = 0
    Timer1.Interval = 1
    SoundMeter.BUFFER_SIZE = 800
    
End Sub

Private Sub Command1_Click()
    SoundMeter.StartInput
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
    SoundMeter.StopInput
End Sub

Private Sub Timer1_Timer()

    ProgressBar1.Value = SoundMeter.getVolume(buffaddress)
    drawScope
    
End Sub

Private Sub drawScope()
   Dim n As Integer
   Dim avg As Integer
    CopyStructFromPtr audbytearray, buffaddress, SoundMeter.BUFFER_SIZE
    
    scopeBox.Cls
    tempval = 0
    avg = 0
    posval = 0
    For n = 0 To SoundMeter.BUFFER_SIZE - 1
        scopeBox.Width = SoundMeter.BUFFER_SIZE - 1
        scopeBox.PSet (n, (scopeBox.Height - _
          (audbytearray.bytes(n) * scopeBox.Height / 255)))
        posval = audbytearray.bytes(n) - 128
        If posval < 0 Then posval = 0 - posval
        If posval > tempval Then tempval = posval
    Next n
End Sub
