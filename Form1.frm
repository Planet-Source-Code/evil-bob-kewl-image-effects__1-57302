VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Evil Bob's Image Effects"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   13920
      TabIndex        =   34
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   7440
      Width           =   1095
   End
   Begin VB.PictureBox imgNo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   6000
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   32
      Top             =   8040
      Width           =   135
   End
   Begin VB.CommandButton cmdPointalism 
      Caption         =   "Pointalism"
      Height          =   375
      Left            =   8160
      TabIndex        =   31
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPixilated 
      Caption         =   "Pixilated"
      Height          =   375
      Left            =   6960
      TabIndex        =   30
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmboss 
      Caption         =   "Emboss"
      Height          =   375
      Left            =   5760
      TabIndex        =   29
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSilk2 
      Caption         =   "Silk2"
      Height          =   375
      Left            =   8160
      TabIndex        =   28
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSilhuette 
      Caption         =   "Silhuette"
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdFlatten 
      Caption         =   "Flatten"
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSilk 
      Caption         =   "Silk"
      Height          =   375
      Left            =   8160
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSunShinnyDay 
      Caption         =   "SunShinnyDay"
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdGlowInTheDark 
      Caption         =   "GlowInDark"
      Height          =   375
      Left            =   5760
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseRGB 
      Caption         =   "DecreaseRGB"
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseBlue 
      Caption         =   "DecreaseBlue"
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseGreen 
      Caption         =   "DecreaseG"
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecreaseRed 
      Caption         =   "DecreaseRed"
      Height          =   375
      Left            =   8160
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncoherence 
      Caption         =   "Incoherence"
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSandBlasting 
      Caption         =   "Sand Blasting"
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdWarped2 
      Caption         =   "Warped2"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseRGB 
      Caption         =   "IncreaseRGB"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseBlue 
      Caption         =   "IncreaseBlue"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseGreen 
      Caption         =   "IncreaseGreen"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncreaseRed 
      Caption         =   "IncreaseRed"
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdWarped 
      Caption         =   "Warped"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoGreen 
      Caption         =   "NoGreen"
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoBlue 
      Caption         =   "NoBlue"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoRed 
      Caption         =   "NoRed"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlur2 
      Caption         =   "Blur2"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlur1 
      Caption         =   "Blur1"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDarken 
      Caption         =   "Darken"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdLighten 
      Caption         =   "Lighten"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNegative 
      Caption         =   "Negative"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdGreyScale 
      Caption         =   "GreyScale"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   9480
      ScaleHeight     =   7185
      ScaleWidth      =   5505
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7185
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   6720
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlur1_Click()
Blur Picture1, Picture2
End Sub

Private Sub cmdBlur2_Click()
Blur2 Picture1, Picture2
End Sub

Private Sub cmdDarken_Click()
Darken Picture1, Picture2, 50
End Sub

Private Sub cmdDecreaseBlue_Click()
DecreaseBlue Picture1, Picture2, 50
End Sub

Private Sub cmdDecreaseGreen_Click()
DecreaseGreen Picture1, Picture2, 50
End Sub

Private Sub cmdDecreaseRed_Click()
DecreaseRed Picture1, Picture2, 50
End Sub

Private Sub cmdDecreaseRGB_Click()
DecreaseRGB Picture1, Picture2, 50, 100, 50
End Sub

Private Sub cmdEmboss_Click()
Emboss Picture1, Picture2, 10
End Sub

Private Sub cmdFlatten_Click()
Flatten Picture1, Picture2
End Sub

Private Sub cmdGlowInTheDark_Click()
GlowInTheDark Picture1, Picture2, 1
End Sub

Private Sub cmdGreyScale_Click()
GrayScale Picture1, Picture2
End Sub

Private Sub cmdIncoherence_Click()
Incoherence Picture1, Picture2
End Sub

Private Sub cmdIncreaseBlue_Click()
IncreaseBlue Picture1, Picture2, 50
End Sub

Private Sub cmdIncreaseGreen_Click()
IncreaseGreen Picture1, Picture2, 50
End Sub

Private Sub cmdIncreaseRed_Click()
IncreaseRed Picture1, Picture2, 50
End Sub

Private Sub cmdIncreaseRGB_Click()
IncreaseRGB Picture1, Picture2, 50, 100, 50
End Sub

Private Sub cmdLighten_Click()
Lighten Picture1, Picture2, 50
End Sub

Private Sub cmdLoadImage_Click()
CMDialog1.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
CMDialog1.Filter = "Files (*.*)|*.*"
CMDialog1.DefaultExt = ""
CMDialog1.DialogTitle = "Open Image"
On Error GoTo No_Open
CMDialog1.ShowOpen
Open CMDialog1.FileName For Input As #1
Picture1.Picture = LoadPicture(CMDialog1.FileName)
Close 1
Exit Sub
No_Open:
Resume ExitLine
ExitLine:
Exit Sub
End Sub

Private Sub cmdNegative_Click()
Negative Picture1, Picture2
End Sub

Private Sub cmdNoBlue_Click()
NoBlue Picture1, Picture2
End Sub

Private Sub cmdNoGreen_Click()
NoGreen Picture1, Picture2
End Sub

Private Sub cmdNoRed_Click()
NoRed Picture1, Picture2
End Sub

Private Sub cmdPixilated_Click()
Pixilated Picture1, Picture2
End Sub

Private Sub cmdPointalism_Click()
Pointalism Picture1, Picture2, 2
End Sub

Private Sub cmdSandBlasting_Click()
SandBlasting Picture1, Picture2
End Sub

Private Sub cmdSave_Click()
CMDialog1.Filter = "Files (*.bmp)|*.bmp"
CMDialog1.DefaultExt = "bmp"
CMDialog1.DialogTitle = "Save Picture File"
CMDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
On Error GoTo No_Save
CMDialog1.ShowSave
Open CMDialog1.FileName For Output As #2
SavePicture Picture2.Image, CMDialog1.FileName
Close 2
Exit Sub
No_Save:
Resume ExitLine
ExitLine:
Exit Sub
End Sub

Private Sub cmdSilhuette_Click()
Silhuette Picture1, Picture2
End Sub

Private Sub cmdSilk_Click()
AmphibiousSilk Picture1, Picture2
End Sub

Private Sub cmdSilk2_Click()
Silk2 Picture1, Picture2
End Sub

Private Sub cmdSunShinnyDay_Click()
GlowInTheDark Picture1, Picture2, 10
End Sub

Private Sub cmdWarped_Click()
Warped Picture1, Picture2
End Sub

Private Sub cmdWarped2_Click()
Warped2 Picture1, Picture2
End Sub

Public Sub Pointalism(PicSrc As PictureBox, PicDest As PictureBox, Radius As Integer)
X = Radius
z = PicDest.Width / 15
Y = Radius
Z2 = PicDest.Height / 15

PicDest.FillStyle = 0
Set PicDest.Picture = imgNo.Image

Do Until X >= z
    
    Y = Radius
    
    Do Until Y >= Z2
        u = u + 1: If u = 4000 Then DoEvents: u = 0
        
        pix = PicSrc.Point(X * 15, Y * 15)
        UnRGB pix, R%, G%, b%
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If b < 0 Then b = 0
        
        PicDest.FillColor = RGB(R, G, b)
        PicDest.Circle (X * 15, Y * 15), Radius * 15, RGB(R, G, b)
        
    Y = Y + (Radius * 2)
    Loop

X = X + (Radius * 2)
Loop
End Sub
