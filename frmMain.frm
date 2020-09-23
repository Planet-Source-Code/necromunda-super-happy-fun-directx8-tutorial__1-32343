VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectX"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmdiFile 
      Left            =   4920
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Picture Files |*.bmp .jpeg|"
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   4920
   End
   Begin VB.ListBox lstReport 
      Height          =   1230
      Left            =   6240
      TabIndex        =   7
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Frame fmeExtra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Extra Options"
      Height          =   2895
      Left            =   6120
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton cmdMoveCam 
         Caption         =   "Move Camera"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdModel 
         Caption         =   "Change Model"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdTexture 
         Caption         =   "Change Texture"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox chkDrag 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Drag Mode"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox chkZ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Z"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   2040
         Width           =   495
      End
      Begin VB.CheckBox chkY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Y"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.CheckBox chkX 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&X"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   2040
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.CheckBox chkMove 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Translate"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chkScale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Scale"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkRotate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Rotate in:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   3000
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin VB.Frame fmeSetup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Setup Options"
      Height          =   1215
      Left            =   6120
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
      Begin VB.ComboBox cmbMode 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   1320
         List            =   "frmMain.frx":0010
         TabIndex        =   15
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optWindow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Window"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fullscreen"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Mode:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Unload Objects and Exit"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "St&art Scene"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picDirectX 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   120
      Picture         =   "frmMain.frx":003C
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   2
      ToolTipText     =   "Click and hold to drag!"
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim FPS_LastCheck As Long
Dim FPS_Count As Long
Dim FPS_Current As Integer



Private Sub Form_Load()
        
    Call Randomize
    cmbMode.ListIndex = 0
    currentModel = "default"
    currentTexture = "default"
        
End Sub
Private Sub Form_Unload(Cancel As Integer)
            
    'Unload all the objects that have been loaded
    Call unloadObjects

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
        'Check for keypresses
    
        Select Case KeyCode
            Case Is = vbKeyEscape
                Call unloadObjects
            Case Is = vbKeyR        'Rotate
                chkRotate.Value = IIf(chkRotate.Value = 1, 0, 1)
            Case Is = vbKeyS        'Scale
                chkScale.Value = IIf(chkScale.Value = 1, 0, 1)
            Case Is = vbKeyT        'Transform
                chkMove.Value = IIf(chkMove.Value = 1, 0, 1)
            Case Is = vbKeyX        'X-Axis Rotate
                chkX.Value = IIf(chkX.Value = 1, 0, 1)
            Case Is = vbKeyY        'Y-Axis Rotate
                chkY.Value = IIf(chkY.Value = 1, 0, 1)
            Case Is = vbKeyZ        'Z-Axis Rotate
                chkZ.Value = IIf(chkZ.Value = 1, 0, 1)
        End Select

End Sub


Private Sub cmdModel_Click()

    'Set the commondialog to display only .X  model files
    cmdiFile.Filter = "Microsoft DirectX .X Files (*.X) |*.x"
    cmdiFile.ShowOpen
    
    'Check if a model has been selected and load the model if it has
    If cmdiFile.FileName <> "" Then Call loadModel(cmdiFile.FileName) '[mModels]
   
End Sub
Private Sub cmdTexture_Click()
    
    'Set filter to display only bitmaps and jpegs
    cmdiFile.Filter = "Picture Files (*.bmp, *.jpeg, *.jpg) | *.bmp; *.jpeg; *.jpg"
    cmdiFile.ShowOpen
    
    'Check for file and load
    If cmdiFile.FileName <> "" Then Call setTexture(cmdiFile.FileName) '[mTextures]

End Sub
Private Sub cmdMoveCam_Click()
    Dim moveTo As D3DVECTOR
    
    On Error GoTo quitsub
    
    'Get the user input for the camera location and
    'reset the camera
    moveTo.X = InputBox("Set X location", "Move Camera")
    moveTo.Y = InputBox("Set Y location", "Move Camera")
    moveTo.Z = InputBox("Set Z location", "Move Camera")
    
    With moveTo
        If .X > 90 Then .X = 90
        If .Y > 90 Then .Y = 90
        If .Z > 90 Then .Z = 90
    End With
    
    Call moveCamera(moveTo) '[mGeometry]

quitsub:
End Sub
Private Sub cmdStart_Click()
                
    lstReport.Clear
    currentApp = notRunning
    
    'initD3D [mInit]
    If InitD3D(optWindow.Value, IIf(optWindow.Value, picDirectX.hWnd, frmFullDisp.hWnd)) = True Then
            
        Set picDirectX = Nothing
        addReport ("Window DX Scene Loaded!")
        addReport ("--------")
        
        cmdTexture.Enabled = True
        cmdMoveCam.Enabled = True
        cmdModel.Enabled = True
                
        Select Case cmbMode.text
            Case Is = "Normal"
                currentApp = Normal
                initGeometry '[mGeometry]
            
            Case Is = "Lighting"
                currentApp = Lighting
                initLightingGeo '[mLights]
                SetupLights '[mLights]
            
            Case Is = "Texturing"
                currentApp = Texturing
                initTexCube '[mTextures]
                
            Case Is = "Modelling"
                currentApp = Modelling
                loadModel (currentModel) '[mModels]
                SetupModelLights '[mModels]
            Case Else
                currentApp = Normal
                initGeometry '[mInit]
            
        End Select
        
        
        initMatrix
        If chkDrag.Value = 0 Then Timer1.Enabled = True
        
        Do While Not currentApp = notRunning
            Call Render '[mRender]
        
            If GetTickCount() - FPS_LastCheck >= 100 Then
                FPS_Current = FPS_Count * 10
                FPS_Count = 0
                FPS_LastCheck = GetTickCount()
            End If
        
            FPS_Count = FPS_Count + 1
            If Not optFull.Value Then Me.Caption = FPS_Current & "fps"
                     
            'Let Windows take a breath
            DoEvents
        
        Loop
    Else
        MsgBox "Initialisation failed!"
    End If
    
    
End Sub
Private Sub cmdExit_Click()

    Call unloadObjects '[mObjects]

End Sub


Private Sub chkDrag_Click()
    
    If chkDrag.Value = 1 Then
        Timer1.Enabled = False
        picDirectX.Enabled = True
    
        chkRotate.Enabled = False
        chkScale.Enabled = False
        chkMove.Enabled = False
        
        chkX.Enabled = False
        chkY.Enabled = False
        chkZ.Enabled = False
    Else
        If currentApp <> notRunning Then Timer1.Enabled = True
        picDirectX.Enabled = False
    
        chkRotate.Enabled = True
        chkScale.Enabled = True
        chkMove.Enabled = True
        
        chkX.Enabled = True
        chkY.Enabled = True
        chkZ.Enabled = True
    End If
        
End Sub
Private Sub chkRotate_Click()
    
    If chkRotate.Value = 1 Then
        chkX.Visible = True: chkY.Visible = True: chkZ.Visible = True
    Else
        chkX.Visible = False: chkY.Visible = False: chkZ.Visible = False
    End If

End Sub

Private Sub optFull_Click()
    
    chkDrag.Enabled = False

    If chkDrag.Value = 1 Then
        chkDrag.Value = 0
        Call chkDrag_Click
    End If

End Sub
Private Sub optWindow_Click()

    chkDrag.Enabled = True

End Sub

Private Sub cmbMode_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub
Private Sub cmbMode_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub
Private Sub cmbMode_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub



Private Sub picDirectX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim matPic As D3DMATRIX
    
    If Button <> 0 Then
        D3DXMatrixRotationYawPitchRoll matPic, _
            (X / picDirectX.ScaleWidth) * 2.5, _
            (Y / picDirectX.ScaleHeight) * 2.5, _
            0
        D3DDevice.SetTransform D3DTS_WORLD, matPic
    End If
    
End Sub


Private Sub Timer1_Timer()
    Dim matWorld As D3DMATRIX
    Dim matRotate As D3DMATRIX
    Dim matScale As D3DMATRIX
    Dim matTrans As D3DMATRIX
 
    If chkScale.Value = 1 Then
        D3DXMatrixScaling matScale, _
             Abs(Sin(Timer)), _
             Abs(Sin(Timer)), _
             Abs(Sin(Timer))
    Else
        D3DXMatrixIdentity matScale
    End If
    
    
    If chkRotate.Value = 1 Then
        D3DXMatrixRotationYawPitchRoll matRotate, _
            IIf(chkX.Value = 1, 1.2 * Sin(Timer), 0), _
            IIf(chkY.Value = 1, 0.8 * Timer, 0), _
            IIf(chkZ.Value = 1, Timer, 0)
    Else
        D3DXMatrixIdentity matRotate
    End If
    
    
    If chkMove.Value = 1 Then
        D3DXMatrixTranslation matTrans, _
            Cos(Timer), _
            Sin(Timer), _
            0
    Else
        D3DXMatrixIdentity matTrans
    End If
            
    D3DXMatrixMultiply matWorld, matRotate, matTrans
    D3DXMatrixMultiply matWorld, matScale, matWorld
             
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
       
End Sub

