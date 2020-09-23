VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PixsxiP Lab "
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   5760
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fmControlsCont 
      BackColor       =   &H00E0E0E0&
      Height          =   6885
      Left            =   6750
      TabIndex        =   5
      Top             =   -75
      Width           =   3855
      Begin VB.CommandButton cmdSamFor 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton cmdSamBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3960
         Width           =   255
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   26
         Text            =   "0"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   25
         Text            =   "0.5"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   24
         Text            =   "1"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   23
         Text            =   "0.15"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   21
         Text            =   "0"
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         Text            =   "0.25"
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   19
         Text            =   "0.75"
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Text            =   "0.15"
         Top             =   1200
         Width           =   500
      End
      Begin VB.CommandButton cmdOpenText 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   270
         Index           =   1
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3480
         Width           =   255
      End
      Begin VB.CommandButton cmdOpenText 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   270
         Index           =   0
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txtTex2 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtTex1 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save Shader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Load Shader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton cmdCompile 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6480
         Width           =   975
      End
      Begin VB.TextBox txtShader 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   4245
         Width           =   3615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Samples"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2350
         TabIndex        =   36
         Top             =   3990
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PixsxiP Lab"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   200
         TabIndex        =   35
         Top             =   240
         Width           =   3465
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   3840
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3840
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2860
         TabIndex        =   34
         Top             =   1950
         Width           =   120
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2020
         TabIndex        =   33
         Top             =   1950
         Width           =   105
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1180
         TabIndex        =   32
         Top             =   1950
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   340
         TabIndex        =   31
         Top             =   1950
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2860
         TabIndex        =   30
         Top             =   1230
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2020
         TabIndex        =   29
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1180
         TabIndex        =   28
         Top             =   1230
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   340
         TabIndex        =   27
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Constant c1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Constant c0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3840
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   3840
         Y1              =   3905
         Y2              =   3905
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   3840
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   3840
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texture 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texture 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixel Shader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   3990
         Width           =   1020
      End
   End
   Begin VB.Frame fmCanvasCont 
      BackColor       =   &H00E0E0E0&
      Height          =   5610
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   6720
      Begin VB.PictureBox Canvas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   5300
         Left            =   100
         ScaleHeight     =   349
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   429
         TabIndex        =   1
         Top             =   200
         Width           =   6500
      End
   End
   Begin VB.Frame fmErrorCont 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   5475
      Width           =   6720
      Begin VB.TextBox txtError 
         BackColor       =   &H00C0C0C0&
         Height          =   900
         Left            =   80
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   6570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DirectX Assembler Result"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   3
         Top             =   120
         Width           =   2130
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This program is not that difficult. Its just coming
'up with the assembly code to make it work. The only reaon I made
'it was for my RPG game so the seams would get covered up,
'but I thought I would share with everyone at PSC.

Private Sub cmdCompile_Click()
 
 If Not HasTexture1 Or Not HasTexture2 Then MsgBox "Please load both textures...", vbInformation, "PixsxiP Lab Information": Exit Sub
 Screen.MousePointer = 11
 If PS_Handle Then
  Call D3DD.DeletePixelShader(PS_Handle)
  PS_Handle = 0
 End If
 
 PixelShader = txtShader.Text
 UpdateShader
 Render
 Screen.MousePointer = 0
 
End Sub

Private Sub cmdLoad_Click()
 Dim tmpString As String
 Dim FSys As New FileSystemObject
 Dim InputStream As TextStream
 
 CD.Filter = "Vertex Shader File (*.txt)|*.txt"
 CD.DialogTitle = "Open Pixel Shader"
 CD.InitDir = App.Path
 CD.FileName = ""
 CD.ShowOpen
 
 If CD.FileName <> "" Then
  txtShader.Text = ""
  Set InputStream = FSys.OpenTextFile(CD.FileName)
  
  Do Until InputStream.AtEndOfStream = True
   tmpString = InputStream.ReadLine
   txtShader.Text = txtShader.Text & tmpString & vbNewLine
  Loop
 End If
 
 Set InputStream = Nothing
 Set FSys = Nothing
End Sub

Private Sub cmdOpenText_Click(Index As Integer)

 CD.Filter = "Bitmap Image (*.bmp)|*.bmp|TGA Image (*.tga)|*.tga|JPG Image (*.jpg)|*.jpg"
 CD.DialogTitle = "Open Texture" & Index
 CD.InitDir = App.Path
 CD.FileName = ""
 CD.ShowOpen
 
 If CD.FileName <> "" Then
  Select Case Index
   Case 0
    Set Mesh.Texture(0) = Nothing
    HasTexture1 = True
    txtTex1.Text = CD.FileTitle
    Set Mesh.Texture(0) = D3DX.CreateTextureFromFile(D3DD, CD.FileName)
   Case 1
    Set Mesh.Texture(1) = Nothing
    HasTexture2 = True
    txtTex2.Text = CD.FileTitle
    Set Mesh.Texture(1) = D3DX.CreateTextureFromFile(D3DD, CD.FileName)
  End Select
 End If
 
End Sub

Private Sub cmdSamBack_Click()
 CycleSamplesBackward
End Sub

Private Sub cmdSamFor_Click()
 CycleSamplesForward
End Sub

Private Sub cmdSave_Click()
 Dim tmpString As String
 Dim FSys As New FileSystemObject
 Dim OutputStream As TextStream
 
 CD.Filter = "Vertex Shader File (*.txt)|*.txt"
 CD.DialogTitle = "Open Pixel Shader"
 CD.FileName = ""
 CD.flags = &HF
 CD.ShowSave
 
 If CD.FileName <> "" Then
  Set OutputStream = FSys.CreateTextFile(CD.FileName)
  OutputStream.Write txtShader.Text
 End If
 
 Set OutputStream = Nothing
 Set FSys = Nothing
End Sub

Private Sub Form_Load()
 PixelShader = txtShader.Text
 SavePicture frmTextures.Water.Picture, "Water.bmp"
 SavePicture frmTextures.Terrain.Picture, "Terrain.bmp"
 MakeSamples
 Initialize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Kill "Water.bmp"
 Kill "Terrain.bmp"
 Unload frmTextures
 Cleanup_DX8
End Sub

Private Sub txtc0_LostFocus(Index As Integer)
 'DUH!!! HOW STUPID CAN I BE? I had clng() instead of csng()...
 Dim tmpString As String
 
 If Left$(txtc0(Index).Text, 1) = "." Then
  tmpString = txtc0(Index).Text
  txtc0(Index).Text = CSng("0" & tmpString)
 ElseIf Left$(txtc0(Index).Text, 2) = "-." Then
  tmpString = Mid$(txtc0(Index).Text, 2, 50)
  txtc0(Index).Text = CSng("-0" & Trim$(tmpString))
 ElseIf CSng(txtc0(Index).Text) > 1 Then
  txtc0(Index).Text = 1
 ElseIf CSng(txtc0(Index).Text) < -1 Then
  txtc0(Index).Text = -1
 End If
 
End Sub

Private Sub txtc1_LostFocus(Index As Integer)
 Dim tmpString As String
 
 If Left$(txtc1(Index).Text, 1) = "." Then
  tmpString = txtc1(Index).Text
  txtc1(Index).Text = CSng("0" & tmpString)
 ElseIf Left$(txtc1(Index).Text, 2) = "-." Then
  tmpString = Mid$(txtc1(Index).Text, 2, 50)
  txtc1(Index).Text = CSng("-0" & Trim$(tmpString))
 ElseIf CSng(txtc1(Index).Text) > 1 Then
  txtc1(Index).Text = 1
 ElseIf CSng(txtc1(Index).Text) < -1 Then
  txtc1(Index).Text = -1
 End If
 
End Sub
