VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Find That Zip"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ThumbsCtl TC1 
      Height          =   5055
      Left            =   30
      TabIndex        =   7
      Top             =   375
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   8916
      BackColor       =   16777215
      ForeColor       =   0
      ImageToolTip    =   0   'False
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selected  Zip File"
      Height          =   315
      Left            =   1995
      TabIndex        =   5
      Top             =   30
      Width           =   2040
   End
   Begin VB.TextBox txtUnzipped 
      Height          =   315
      Left            =   4140
      TabIndex        =   1
      Top             =   300
      Width           =   3390
   End
   Begin VB.CommandButton cmdUnzip 
      Caption         =   "Unzip To"
      Height          =   330
      Left            =   7590
      TabIndex        =   0
      Top             =   300
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4140
      TabIndex        =   8
      Top             =   735
      Width           =   4425
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File Count ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   75
      TabIndex        =   6
      Top             =   30
      Width           =   1890
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Resize large pictures to fit and save space.       (width < 300 pixels)"
      Height          =   405
      Left            =   75
      TabIndex        =   4
      Top             =   5970
      Width           =   3315
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Zip filename and picture filename must be the    same and placed in DataFolder."
      Height          =   390
      Left            =   75
      TabIndex        =   3
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Folder where unzipped file is located"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4155
      TabIndex        =   2
      Top             =   75
      Width           =   3210
   End
   Begin VB.Image imgDisplay 
      BorderStyle     =   1  'Fixed Single
      Height          =   4305
      Left            =   4140
      Top             =   1080
      Width           =   4590
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   45
      Top             =   5475
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'*
'*         Project Name : Find That Zip
'*        Version Number: 1.0.2
'*           Author Name: Ken Foster
'*                 Date : February 04, 2007
'*        Freeware - Use anyway you want.
'*
'****************************************************
'***************** Table of Procedures *************
'   Private Sub Form_Load
'   Private Sub TC1_ImageClick
'   Private Sub cmdUnzip_Click
'   Private Sub cmdDelete_Click
'   Private Sub Form_Unload
'   Private Function extARchive
'***************** End of Table ********************

'you will need three files once project is compiled into an exe.
'place all three into the same folder
'the .exe file
'DataFolder        ...this will be a subfolder
'unzip32.dll
' I preferred this method over a database because I still have access to the zip file.
' Plus it makes adding updated zipfiles easier.

Option Explicit

Dim stgName As String
Private bzip As CGUnzipFiles

Private Sub Form_Load()
    Set bzip = New CGUnzipFiles
    TC1.Clear
    Call TC1.ScanPath("D:\AProjectFolder0\FindThatProject\DataFolder")   'load images of zip files
    Label6.Caption = "File Count = " & TC1.ListCount                     'load file count
    Screen.MousePointer = 0
End Sub

Private Sub TC1_ImageClick(PicIndex As Integer, FileName As String)
    
    stgName = Right$(FileName, Len(FileName) - InStrRev(FileName, "\"))   'get filename
    stgName = Left$(stgName, Len(stgName) - 4)                  'filename without extension
    Label1.Caption = stgName
    
    'load the picture and zip file
    If FileExists(App.Path & "\DataFolder\" & stgName & ".jpg") Then
        imgDisplay.Picture = LoadPicture(App.Path & "\DataFolder\" & stgName & ".jpg")
    End If
    If FileExists(App.Path & "\DataFolder\" & stgName & ".bmp") Then
        imgDisplay.Picture = LoadPicture(App.Path & "\DataFolder\" & stgName & ".bmp")
    End If
    If FileExists(App.Path & "\DataFolder\" & stgName & ".gif") Then
        imgDisplay.Picture = LoadPicture(App.Path & "\DataFolder\" & stgName & ".gif")
    End If
End Sub

Private Sub cmdUnzip_Click()
    txtUnzipped.Text = GetBrowseDirectory(Form1)
    If txtUnzipped.Text = "" Then Exit Sub
    extARchive App.Path & "\DataFolder\" & stgName & ".zip", txtUnzipped.Text
    MsgBox "File was successfully unzipped!!!"
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure you want to delete this file?", vbYesNo + vbExclamation, "Delete This File!") = vbNo Then Exit Sub
    DeleteFile (App.Path & "\DataFolder\" & stgName & ".zip")
    If FileExists(App.Path & "\DataFolder\" & stgName & ".jpg") Then
       DeleteFile (App.Path & "\DataFolder\" & stgName & ".jpg")
    End If
    If FileExists(App.Path & "\DataFolder\" & stgName & ".bmp") Then
       DeleteFile (App.Path & "\DataFolder\" & stgName & ".bmp")
    End If
    If FileExists(App.Path & "\DataFolder\" & stgName & ".gif") Then
       DeleteFile (App.Path & "\DataFolder\" & stgName & ".gif")
    End If
    
    TC1.Clear
    Call TC1.ScanPath("D:\AProjectFolder0\FindThatProject\DataFolder")   'refresh images
    Label6.Caption = "File Count = " & TC1.ListCount
    Label1.Caption = ""
    imgDisplay.Picture = LoadPicture()                            'clear picture
    MsgBox stgName & " successfully deleted!!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set bzip = Nothing
End Sub

Private Function extARchive(aPath As String, extPath As String)
    With bzip
        .Unzip aPath, extPath
    End With
End Function
