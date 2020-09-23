Attribute VB_Name = "Shader"
Option Explicit

'I'll do a rough explaination.(read DX SDK Doc to make sure on this stuff)
'There is a ton more functions than in the examples, look at
'www.ati.com/developer I believe the have info on this.
'
'The PS_Constants as for the c0 and c1 Registers.
' - c0 = Color 0
' - c1 = Color 1
'
'ps.1.0 is the version your going to use. It goes up to 1.4
' -My GeForce4 can only handle up to 1.3(In REF mode) I think only Radeons 8xxx can handle 1.4? not sure.
' -Thats the downfall
'
'tex t0 and tex t1
' -These are to define the textures
'
'mov (arg)
' -moves the argument to the register
' -Exmaple (mov r0, t0) = move texture 0 to Register 0
'
'dp3
' -Dot Product 3
'
'lrp
' -Linear Interpolation (or however you spell it)
'

Public Sub SetupShaderMain()
 PS_Constants(0, 0) = CSng(frmMain.txtc0(0).Text)
 PS_Constants(1, 0) = CSng(frmMain.txtc0(1).Text)
 PS_Constants(2, 0) = CSng(frmMain.txtc0(2).Text)
 PS_Constants(3, 0) = CSng(frmMain.txtc0(3).Text)
 PS_Constants(0, 1) = CSng(frmMain.txtc1(0).Text)
 PS_Constants(1, 1) = CSng(frmMain.txtc1(1).Text)
 PS_Constants(2, 1) = CSng(frmMain.txtc1(2).Text)
 PS_Constants(3, 1) = CSng(frmMain.txtc1(3).Text)
 
 Set DX_VB = D3DD.CreateVertexBuffer(4 * Len(Mesh.Vertices(0)), D3DUSAGE_WRITEONLY, FVF_ShaderVertex, D3DPOOL_MANAGED)
 PositionMesh 10, 10
 CreateShaderFromCode
End Sub

Public Sub UpdateShader()
 PS_Constants(0, 0) = CSng(frmMain.txtc0(0).Text)
 PS_Constants(1, 0) = CSng(frmMain.txtc0(1).Text)
 PS_Constants(2, 0) = CSng(frmMain.txtc0(2).Text)
 PS_Constants(3, 0) = CSng(frmMain.txtc0(3).Text)
 PS_Constants(0, 1) = CSng(frmMain.txtc1(0).Text)
 PS_Constants(1, 1) = CSng(frmMain.txtc1(1).Text)
 PS_Constants(2, 1) = CSng(frmMain.txtc1(2).Text)
 PS_Constants(3, 1) = CSng(frmMain.txtc1(3).Text)
 
 CreateShaderFromCode
End Sub

Public Sub CreateShaderFromCode()
 On Local Error Resume Next
 Dim ShaderCode As D3DXBuffer
 Dim RetError As String
 Dim Arr() As Long
 Dim Size As Long
 Dim i As Long

 Set ShaderCode = D3DX.AssembleShader(PixelShader, 0, Nothing)
 Size = ShaderCode.GetBufferSize() / 4
 ReDim Arr(Size - 1)
 Call D3DX.BufferGetData(ShaderCode, 0, 4, Size, Arr(0))
 PS_Handle = D3DD.CreatePixelShader(Arr(0))
 
 If PS_Handle Then
  frmMain.txtError.Text = ""
  frmMain.txtError.Text = "Compiled Successfully"
 Else
  frmMain.txtError.Text = ""
  frmMain.txtError.Text = "Error..."
 End If
 If PixelShader = "" Then
  frmMain.txtError.Text = ""
  frmMain.txtError.Text = "No Pixel Shader Defined..."
 End If
 
 Set ShaderCode = Nothing
 DoEvents
 
 D3DD.SetStreamSource 0, DX_VB, Len(Mesh.Vertices(0))
 D3DD.SetVertexShader FVF_ShaderVertex
 D3DD.SetPixelShaderConstant 0, PS_Constants(0, 0), 2
End Sub

'Here down are just some samples.

Public Sub MakeSamples()
 'I set this up for readability
 
 Samples(0) = ";------------------------------------" & vbNewLine & _
              "; Show Texture 0 Version 1.0         " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "mov r0, t0                           "
              
 Samples(1) = ";------------------------------------" & vbNewLine & _
              "; Show Texture 1 Version 1.0         " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r0, t1                           "
              
 Samples(2) = ";------------------------------------" & vbNewLine & _
              "; Gray Scale Texture 0 Version 1.0   " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r0, t0                           " & vbNewLine & _
              "dp3 r0, r0, c0                       "
              
 Samples(3) = ";------------------------------------" & vbNewLine & _
              "; Gray Scale Texture 1 Version 1.0   " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r0, t1                           " & vbNewLine & _
              "dp3 r0, r0, c0                       "
              
 Samples(4) = ";------------------------------------" & vbNewLine & _
              "; Blend_V Textures 1 Version 1.0     " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t1                           " & vbNewLine & _
              "lrp r0, v1, r1, t0                   "
              
 Samples(5) = ";------------------------------------" & vbNewLine & _
              "; Blend_V Textures 2 Version 1.0     " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t0                           " & vbNewLine & _
              "lrp r0, v1, r1, t1                   "
              
 Samples(6) = ";------------------------------------" & vbNewLine & _
              "; Blend_H Textures 1 Version 1.0     " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t1                           " & vbNewLine & _
              "lrp r0, v0, r1, t0                   "
              
 Samples(7) = ";------------------------------------" & vbNewLine & _
              "; Blend_H Textures 2 Version 1.0     " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t0                           " & vbNewLine & _
              "lrp r0, v0, r1, t1                   "
              
 Samples(8) = ";-------------------------------------------" & vbNewLine & _
              "; Gray Scale Blend_V Textures 1 Version 1.0 " & vbNewLine & _
              ";-------------------------------------------" & vbNewLine & _
              "ps.1.0                                      " & vbNewLine & _
              "                                            " & vbNewLine & _
              "tex t0                                      " & vbNewLine & _
              "tex t1                                      " & vbNewLine & _
              "mov r1, t1                                  " & vbNewLine & _
              "lrp r0, v1, r1, t0                          " & vbNewLine & _
              "dp3 r0, r0, c0                              "
              
 Samples(9) = ";-------------------------------------------" & vbNewLine & _
              "; Gray Scale Blend_V Textures 2 Version 1.0 " & vbNewLine & _
              ";-------------------------------------------" & vbNewLine & _
              "ps.1.0                                      " & vbNewLine & _
              "                                            " & vbNewLine & _
              "tex t0                                      " & vbNewLine & _
              "tex t1                                      " & vbNewLine & _
              "mov r1, t0                                  " & vbNewLine & _
              "lrp r0, v1, r1, t1                          " & vbNewLine & _
              "dp3 r0, r0, c0                              "
              
 Samples(10) = ";-------------------------------------------" & vbNewLine & _
               "; Gray Scale Blend_H Textures 1 Version 1.0 " & vbNewLine & _
               ";-------------------------------------------" & vbNewLine & _
               "ps.1.0                                      " & vbNewLine & _
               "                                            " & vbNewLine & _
               "tex t0                                      " & vbNewLine & _
               "tex t1                                      " & vbNewLine & _
               "mov r1, t1                                  " & vbNewLine & _
               "lrp r0, v0, r1, t0                          " & vbNewLine & _
               "dp3 r0, r0, c0                              "
              
 Samples(11) = ";-------------------------------------------" & vbNewLine & _
               "; Gray Scale Blend_H Textures 2 Version 1.0 " & vbNewLine & _
               ";-------------------------------------------" & vbNewLine & _
               "ps.1.0                                      " & vbNewLine & _
               "                                            " & vbNewLine & _
               "tex t0                                      " & vbNewLine & _
               "tex t1                                      " & vbNewLine & _
               "mov r1, t0                                  " & vbNewLine & _
               "lrp r0, v0, r1, t1                          " & vbNewLine & _
               "dp3 r0, r0, c0                              "
              
End Sub

Public Sub CycleSamplesForward()
 
 frmMain.cmdSamBack.Enabled = True
 If SamPosition >= 11 Then SamPosition = 11
 If SamPosition <= 0 Then SamPosition = 0
 If SamPosition >= 11 Then frmMain.cmdSamFor.Enabled = False Else frmMain.cmdSamFor.Enabled = True
 frmMain.txtShader.Text = ""
 frmMain.txtShader.Text = Samples(SamPosition)
 SamPosition = SamPosition + 1
 
End Sub

Public Sub CycleSamplesBackward()
 
 frmMain.cmdSamFor.Enabled = True
 If SamPosition >= 11 Then SamPosition = 11
 If SamPosition <= 0 Then SamPosition = 0
 If SamPosition <= 0 Then frmMain.cmdSamBack.Enabled = False Else frmMain.cmdSamBack.Enabled = True
 frmMain.txtShader.Text = ""
 frmMain.txtShader.Text = Samples(SamPosition)
 SamPosition = SamPosition - 1
 
End Sub
