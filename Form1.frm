VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMAin 
   Caption         =   "File Date Changer"
   ClientHeight    =   4920
   ClientLeft      =   5445
   ClientTop       =   4320
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   Begin VB.CommandButton cmdCdate 
      Caption         =   "Change Date"
      Height          =   375
      Left            =   3150
      TabIndex        =   15
      Top             =   4320
      Width           =   2085
   End
   Begin MSComCtl2.DTPicker dtA 
      Height          =   285
      Left            =   3150
      TabIndex        =   12
      Top             =   2610
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   37580
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   2985
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   2985
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   90
      TabIndex        =   0
      Top             =   2160
      Width           =   2985
   End
   Begin MSComCtl2.DTPicker dtM 
      Height          =   285
      Left            =   3150
      TabIndex        =   13
      Top             =   3240
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   37580
   End
   Begin MSComCtl2.DTPicker dtC 
      Height          =   285
      Left            =   3150
      TabIndex        =   14
      Top             =   3870
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   37580
   End
   Begin VB.Label Label6 
      Caption         =   "New Created Date"
      Height          =   195
      Left            =   3150
      TabIndex        =   11
      Top             =   3600
      Width           =   3165
   End
   Begin VB.Label Label4 
      Caption         =   "New Modify Date"
      Height          =   195
      Left            =   3150
      TabIndex        =   10
      Top             =   2970
      Width           =   3165
   End
   Begin VB.Label Label2 
      Caption         =   "New Access Date"
      Height          =   195
      Left            =   3150
      TabIndex        =   9
      Top             =   2340
      Width           =   3165
   End
   Begin VB.Label lblcdate 
      Height          =   195
      Left            =   3150
      TabIndex        =   8
      Top             =   1530
      Width           =   3165
   End
   Begin VB.Label Label5 
      Caption         =   "Created Date"
      Height          =   195
      Left            =   3150
      TabIndex        =   7
      Top             =   1260
      Width           =   3165
   End
   Begin VB.Label lblmdate 
      Height          =   195
      Left            =   3150
      TabIndex        =   6
      Top             =   990
      Width           =   3165
   End
   Begin VB.Label Label3 
      Caption         =   "Modify Date"
      Height          =   195
      Left            =   3150
      TabIndex        =   5
      Top             =   720
      Width           =   3165
   End
   Begin VB.Label lbladate 
      Height          =   195
      Left            =   3150
      TabIndex        =   4
      Top             =   450
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "Access Date"
      Height          =   195
      Left            =   3150
      TabIndex        =   3
      Top             =   180
      Width           =   3165
   End
End
Attribute VB_Name = "frmMAin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCdate_Click()
SetFileDates Dir1.Path & "\" & File1.FileName, CDate(dtC.Value) + 1, CDate(dtA.Value) + 1, CDate(dtM.Value) + 1
lblcdate.Caption = GetFileCreatedDate(Dir1.Path & "\" & File1.FileName)
lbladate.Caption = GetFileAccessDate(Dir1.Path & "\" & File1.FileName)
lblmdate.Caption = GetFileModifyDate(Dir1.Path & "\" & File1.FileName)
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
lblcdate.Caption = GetFileCreatedDate(Dir1.Path & "\" & File1.FileName)
lbladate.Caption = GetFileAccessDate(Dir1.Path & "\" & File1.FileName)
lblmdate.Caption = GetFileModifyDate(Dir1.Path & "\" & File1.FileName)
End Sub
