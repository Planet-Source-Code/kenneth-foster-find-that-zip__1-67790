VERSION 5.00
Begin VB.UserControl ThumbsCtl 
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   LockControls    =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   4890
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      FillColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   360
      ScaleHeight     =   1260
      ScaleWidth      =   1440
      TabIndex        =   4
      Top             =   3465
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.VScrollBar scrThumbs 
      Height          =   3300
      Left            =   4455
      TabIndex        =   1
      Top             =   90
      Width           =   285
   End
   Begin VB.PictureBox picFundo 
      BackColor       =   &H00FFFFFF&
      Height          =   3300
      Left            =   270
      ScaleHeight     =   3240
      ScaleWidth      =   4140
      TabIndex        =   0
      Top             =   90
      Width           =   4200
      Begin VB.PictureBox picInterior 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3120
         Left            =   45
         ScaleHeight     =   3120
         ScaleWidth      =   4020
         TabIndex        =   2
         Top             =   45
         Width           =   4020
         Begin VB.PictureBox imgFigura 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            Height          =   1500
            Index           =   0
            Left            =   45
            ScaleHeight     =   1500
            ScaleWidth      =   1635
            TabIndex        =   5
            Top             =   45
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblFigura 
            BackColor       =   &H00C0C0C0&
            Height          =   1680
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   -90
            Visible         =   0   'False
            Width           =   1725
         End
      End
   End
   Begin VB.Image imgTeste2 
      Height          =   1500
      Left            =   1980
      Stretch         =   -1  'True
      Top             =   3420
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "ThumbsCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

Const STRETCHMODE = vbPaletteModeNone   'You can find other modes in the "PaletteModeConstants" section of your Object Browser

Private qt_Thumbs() As String
Private imgSize() As String
Private idxImagem As Integer
Private imgPath As String
Private id_BackColor As OLE_COLOR
Private id_ForeColor As OLE_COLOR
Private nr_Columns As Integer
'Private S As cPicScroll
Private ToolTipImg As Boolean

Event ImageClick(PicIndex As Integer, FileName As String)
Event ImageDblClick(PicIndex As Integer, FileName As String)
Event ImageMouseDown(PicIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, FileName As String)
Event ImageMouseMove(PicIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, FileName As String)
Event ImageMouseUp(PicIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, FileName As String)

Enum EnumColumnType
     Auto = 0
     Manual = 1
End Enum
Dim id_ColumnType As EnumColumnType


Public Function AddItem(Caminho As String)

ReDim Preserve qt_Thumbs(UBound(qt_Thumbs) + 1)

qt_Thumbs(UBound(qt_Thumbs)) = Caminho

Call ReorganizaThumbs

End Function


Public Property Let BackColor(cColor As OLE_COLOR)
 
 id_BackColor = cColor
 PropertyChanged "BackColor"
 
 picInterior.BackColor = id_BackColor
 picFundo.BackColor = id_BackColor
  
End Property

Public Property Get BackColor() As OLE_COLOR

   BackColor = picFundo.BackColor

End Property


Public Sub Clear()

 Dim i As Integer
 
 For i = 1 To imgFigura.UBound
     Set imgFigura(i) = Nothing
     Unload imgFigura(i)
     Unload lblFigura(i)
 Next
 
 ReDim qt_Thumbs(0)
 
 picInterior.Width = picFundo.Width
 picInterior.Height = picFundo.Height

 Call SetScrolls(scrThumbs, picInterior.Height - picFundo.Height + 200, 80, picInterior.Height / 3)

End Sub

Public Property Let Columns(newColumns As Integer)

  nr_Columns = newColumns

End Property

Public Property Get Columns() As Integer

  Columns = nr_Columns

End Property

Public Property Get ImageToolTip() As Boolean

  ImageToolTip = ToolTipImg

End Property



Public Property Let ColumnType(NewColumnType As EnumColumnType)

id_ColumnType = NewColumnType
PropertyChanged "ColumnType"

End Property

Public Property Get ColumnType() As EnumColumnType

      ColumnType = id_ColumnType
    
End Property

Public Property Let ForeColor(cColor As OLE_COLOR)

 Dim i As Integer

 id_ForeColor = cColor
 PropertyChanged "ForeColor"
 
 For i = 0 To lblFigura.UBound
   lblFigura(i).BackColor = id_ForeColor
 Next
 
 
End Property

Public Property Let ImageToolTip(newToolTipImg As Boolean)

ToolTipImg = newToolTipImg
 
Dim i As Integer
Dim Texto As String

If ToolTipImg = True Then
   For i = 1 To UBound(qt_Thumbs)
       If qt_Thumbs(i) = "" Then Exit For
       Texto = PegaNomeArquivo(qt_Thumbs(i)) & " (" & Fix(FileLen(qt_Thumbs(i)) / 1024) & "k) - " & imgSize(i)
       imgFigura(i).ToolTipText = Texto
       lblFigura(i).ToolTipText = Texto
   Next i
Else
   For i = 1 To UBound(qt_Thumbs)
       imgFigura(i).ToolTipText = ""
       lblFigura(i).ToolTipText = ""
   Next i

End If
 
End Property

Public Property Get ForeColor() As OLE_COLOR

   ForeColor = lblFigura(0).BackColor

End Property


Private Function PegaNomeArquivo(Path As String) As String
 
Dim nr_UltimaPosicao As Integer
Dim nr_UltimaPosicaoTmp As Integer

nr_UltimaPosicao = 1
nr_UltimaPosicaoTmp = 1

Do While Not nr_UltimaPosicaoTmp = 0
   nr_UltimaPosicaoTmp = InStr(nr_UltimaPosicaoTmp, Path, "\")
   If nr_UltimaPosicaoTmp = 0 Then Exit Do
   nr_UltimaPosicao = nr_UltimaPosicaoTmp
   nr_UltimaPosicaoTmp = nr_UltimaPosicaoTmp + 1
Loop

PegaNomeArquivo = Right(Path, Len(Path) - nr_UltimaPosicao)

End Function

Public Sub RemoveItem(Index As Integer)

 Dim i As Integer
 Dim z As Integer
 Dim ArrayTmp() As String
 
 Dim tmpArraySize As Integer
 
 tmpArraySize = UBound(qt_Thumbs)
 
 For i = 0 To tmpArraySize
     If i = Index Then
        For z = i To tmpArraySize
            If z = tmpArraySize Then
               ReDim qt_Thumbs(UBound(ArrayTmp))
               Call Clear
               qt_Thumbs = ArrayTmp
               Call ReorganizaThumbs
               Exit Sub
            End If
            ReDim Preserve ArrayTmp(z)
            ArrayTmp(z) = qt_Thumbs(z + 1)
        Next z
     End If
     ReDim Preserve ArrayTmp(i)
     ArrayTmp(i) = qt_Thumbs(i)
 Next i

End Sub

Private Sub ReorganizaThumbs()

'Set S = New cPicScroll

Dim Linhas As Integer
Dim Colunas As Integer
Dim qt_Colunas As Integer
Dim Texto As String

If UBound(qt_Thumbs) = 0 Then Exit Sub


If id_ColumnType = Auto Then
   qt_Colunas = Fix(picInterior.Width / (1770 + 65))
Else
   qt_Colunas = nr_Columns
End If

If qt_Colunas <= 0 Then Exit Sub

Screen.MousePointer = 11

For Linhas = 0 To UBound(qt_Thumbs) Step qt_Colunas

    For Colunas = 1 To qt_Colunas
         
        If (Linhas + Colunas) > UBound(qt_Thumbs) Then GoTo AcertaScroll
         
        If imgFigura.UBound < (Linhas + Colunas) Then
           Load imgFigura(Linhas + Colunas)
           Load lblFigura(Linhas + Colunas)
           ReDim Preserve imgSize(Linhas + Colunas)
           
           picOriginal.Picture = LoadPicture(qt_Thumbs(Linhas + Colunas))
           imgTeste2.Picture = picOriginal.Picture
           imgTeste2.Refresh
           imgSize(Linhas + Colunas) = Fix(imgTeste2.Picture.Width / 26.4586148648649) & " x " & Fix(imgTeste2.Picture.Height / 26.4586148648649)
           
           Call Tamanho_Image(Linhas + Colunas, 1500, 1635)
           Call ReSize(imgTeste2.Width, imgTeste2.Height, Linhas + Colunas)
        
           lblFigura(Linhas + Colunas).Left = ((1770 + 65) * (Colunas - 1))
           lblFigura(Linhas + Colunas).Top = ((1590 + 200) * (Linhas / qt_Colunas)) + 200
           lblFigura(Linhas + Colunas).Visible = True
            
           imgFigura(Linhas + Colunas).Left = lblFigura(Linhas + Colunas).Left + ((lblFigura(Linhas + Colunas).Width - imgFigura(Linhas + Colunas).Width) / 2)             'imgFigura(Linhas + Colunas).Left + ((1590 / 2) - (imgFigura(Linhas + Colunas).Width / 2))
           imgFigura(Linhas + Colunas).Top = lblFigura(Linhas + Colunas).Top + ((lblFigura(Linhas + Colunas).Height - imgFigura(Linhas + Colunas).Height) / 2)                       'imgFigura(Linhas + Colunas).Top + ((lblFigura(0).Height / 2) - (imgFigura(Linhas + Colunas).Height / 2))
           imgFigura(Linhas + Colunas).Visible = True
               
           If ToolTipImg = True Then
              Texto = PegaNomeArquivo(qt_Thumbs(Linhas + Colunas)) & " (" & Fix(FileLen(qt_Thumbs(Linhas + Colunas)) / 1024) & "k) - " & imgSize(Linhas + Colunas)
              imgFigura(Linhas + Colunas).ToolTipText = Texto
              lblFigura(Linhas + Colunas).ToolTipText = Texto
           Else
              imgFigura(Linhas + Colunas).ToolTipText = ""
              lblFigura(Linhas + Colunas).ToolTipText = ""
           End If
               
           DoEvents
        End If
        
    
    Next Colunas

Next Linhas

Screen.MousePointer = 0

AcertaScroll:

If UBound(qt_Thumbs) Mod qt_Colunas > 0 Then
   Linhas = (UBound(qt_Thumbs) \ qt_Colunas) + 1
Else
   Linhas = (UBound(qt_Thumbs) \ qt_Colunas)
End If

picInterior.Height = ((lblFigura(0).Height + 130) * (Linhas)) + 200
   
'Call S.SetUpScrollBars(Nothing, scrThumbs, picInterior.Height, picInterior.Height / 50, picInterior.Height / 10, picInterior.Width - picFundo.Width, 40, picInterior.Width / 2)

Screen.MousePointer = 0


End Sub


Private Function ReSize(Xw As Integer, Yw As Integer, Indice As Integer)
On Error GoTo fout

imgFigura(Indice).Width = Xw
imgFigura(Indice).Height = Yw
'picReduzido
imgFigura(Indice).Cls

imgFigura(Indice).PaintPicture imgTeste2.Picture, 0, 0, imgFigura(Indice).Width, imgFigura(Indice).Height, 0, 0, picOriginal.Width, picOriginal.Height

imgFigura(Indice).Refresh

'CenterControl pic, frmMain
Exit Function

fout:
If Err.Number = 6 Then
    MsgBox "This Picture is to big.", vbCritical, App.Title
End If
End Function


Public Sub ScanPath(Path As String)

Dim MyPath As String
Dim MyName As String

If Right(Path, 1) <> "\" Then
   Path = Path & "\"
End If

MyPath = Path
MyName = Dir(MyPath, vbDirectory)
Do While MyName <> ""
    If MyName <> "." And MyName <> ".." Then
       If UCase(Right(MyName, 4)) = ".JPG" Or UCase(Right(MyName, 4)) = ".GIF" Or UCase(Right(MyName, 4)) = ".BMP" Then
          Call AddItem(MyPath & MyName)
       End If
    End If
    MyName = Dir
Loop

End Sub

Private Sub Tamanho_Image(Index As Integer, Height As Integer, Width As Integer)

    imgTeste2.Visible = False


    If imgTeste2.Picture Then
        imgTeste2.Height = imgTeste2.Picture.Height
        imgTeste2.Width = imgTeste2.Picture.Width


        If imgTeste2.Picture.Height > imgTeste2.Picture.Width Then
            imgTeste2.Height = Height
            imgTeste2.Width = imgTeste2.Width / (imgTeste2.Picture.Height / imgTeste2.Height)


            If imgTeste2.Width > Width Then
                imgTeste2.Width = Width
                imgTeste2.Height = imgTeste2.Picture.Height / (imgTeste2.Picture.Width / imgTeste2.Width)
            End If

        End If


        If imgTeste2.Picture.Width > imgTeste2.Picture.Height Then
            imgTeste2.Width = Width
            imgTeste2.Height = imgTeste2.Height / (imgTeste2.Picture.Width / imgTeste2.Width)


            If imgTeste2.Height > Height Then
                imgTeste2.Height = Height
                imgTeste2.Width = imgTeste2.Picture.Width / (imgTeste2.Picture.Height / imgTeste2.Height)
            End If

        End If

        If imgTeste2.Picture.Width = imgTeste2.Picture.Height Then
            imgTeste2.Height = Height
            imgTeste2.Width = imgTeste2.Width / (imgTeste2.Picture.Height / imgTeste2.Height)
        End If

    End If

End Sub


Private Sub imgFigura_Click(Index As Integer)

idxImagem = Index

RaiseEvent ImageClick(Index, qt_Thumbs(Index))

End Sub

Private Sub imgFigura_DblClick(Index As Integer)
idxImagem = Index

RaiseEvent ImageDblClick(Index, qt_Thumbs(Index))

End Sub


Private Sub imgFigura_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ds_Imagem As String
ds_Imagem = qt_Thumbs(Index)

RaiseEvent ImageMouseDown(Index, Button, Shift, X, Y, ds_Imagem)

End Sub

Private Sub imgFigura_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ds_Imagem As String
ds_Imagem = qt_Thumbs(Index)

RaiseEvent ImageMouseMove(Index, Button, Shift, X, Y, ds_Imagem)

End Sub


Private Sub imgFigura_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ds_Imagem As String
ds_Imagem = qt_Thumbs(Index)

RaiseEvent ImageMouseUp(Index, Button, Shift, X, Y, ds_Imagem)

End Sub


Private Sub scrThumbs_Change()
    MoveVScroll scrThumbs, picInterior
    DoEvents
    picInterior.SetFocus
End Sub


Private Sub scrThumbs_Scroll()
    MoveVScroll scrThumbs, picInterior
    DoEvents
    picInterior.SetFocus
End Sub


Private Sub UserControl_Initialize()


ReDim qt_Thumbs(0)

nr_Columns = 1

Call ReorganizaThumbs


End Sub

Private Sub MoveVScroll(V_SB As VScrollBar, pic As PictureBox)
    If (((pic.Height / V_SB.Max) * V_SB.Value) * -1) >= (picFundo.Height * -1) Then
       pic.Top = 0
    Else
       pic.Top = (((pic.Height / V_SB.Max) * V_SB.Value) * -1) + picFundo.Height
    End If
End Sub
Private Sub SetScrolls(V_SB As VScrollBar, lVer_MaxLen As Long, iVer_SmallChange As Integer, lVer_LargeChange As Long)
    
    If lVer_MaxLen < 0 Then lVer_MaxLen = 0
    If lVer_MaxLen < 32000 Then
       V_SB.Max = lVer_MaxLen
    End If
    V_SB.SmallChange = iVer_SmallChange
    V_SB.LargeChange = lVer_LargeChange

End Sub


Public Property Get ListCount() As Integer
    ListCount = UBound(qt_Thumbs)
End Property

Public Property Get ItemData(Index As Integer) As String
    ItemData = qt_Thumbs(Index)
End Property


Public Property Get Path() As String
    Path = imgPath
End Property

Public Property Let Path(sPath As String)
    imgPath = sPath
    PropertyChanged "Path"
End Property



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

id_BackColor = PropBag.ReadProperty("BackColor", &H808080)
picInterior.BackColor = id_BackColor
picFundo.BackColor = id_BackColor
  
id_ForeColor = PropBag.ReadProperty("ForeColor", &HC0C0C0)
lblFigura(0).BackColor = vbWhite 'id_ForeColor

imgPath = PropBag.ReadProperty("Path", "")

nr_Columns = PropBag.ReadProperty("Columns", 1)

id_ColumnType = PropBag.ReadProperty("ColumnType", 0)

ToolTipImg = PropBag.ReadProperty("ImageToolTip", True)

End Sub

Private Sub UserControl_Resize()

picFundo.Top = 10
picFundo.Left = 10
picInterior.Top = 10
picInterior.Left = 10

If (UserControl.Width - scrThumbs.Width - 20) < 0 Then Exit Sub

picFundo.Width = UserControl.Width - scrThumbs.Width - 20
picFundo.Height = UserControl.Height - 20

If picFundo.Width - 20 < 0 Then Exit Sub

picInterior.Width = picFundo.Width - 120
picInterior.Height = picFundo.Height - 120

scrThumbs.Top = 10
scrThumbs.Left = picFundo.Width
scrThumbs.Height = picFundo.Height

Call ReorganizaThumbs

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

 Call PropBag.WriteProperty("BackColor", id_BackColor, &H808080)
 Call PropBag.WriteProperty("ForeColor", id_ForeColor, &HC0C0C0)
 Call PropBag.WriteProperty("Path", imgPath, "")
 Call PropBag.WriteProperty("Columns", nr_Columns, 1)
 Call PropBag.WriteProperty("ColumnType", id_ColumnType, 0)
 Call PropBag.WriteProperty("ImageToolTip", ToolTipImg, True)

End Sub


