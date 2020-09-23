VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Matrix Blender"
   ClientHeight    =   8415
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13965
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   13965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   55
      Top             =   840
      Width           =   1575
   End
   Begin VB.Timer tmrTitle 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "STOP!"
      Height          =   375
      Left            =   12360
      TabIndex        =   40
      Top             =   3840
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   12720
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontBold        =   -1  'True
      FontSize        =   9
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Matrix Field"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbldesc(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbldesc(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbldesc(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbldesc(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbldesc(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbldesc(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbldesc(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbldesc(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbldesc(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbldesc(11)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "pbrMatrixY"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Pic1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "hsbHSpacing"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "hsbVSpacing"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "hsbHNum"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "hsbStartingY"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "hsbLineLength"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtSize"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "hsbDistortion"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkOrg"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "hsbNumOfLines"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbFonts"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "picMatColor"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "chkBoldFont"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkBinary"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkHighLast"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "pbrMatrixX"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdMakeMatrix"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdNext"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Overlaying Image"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lbldesc(9)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbldesc(10)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "imgDefault"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbldesc(13)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lbldesc(14)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbldesc(15)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label3"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Pic2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdSelTransColor"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "picTrans"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdLoadPic"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdInsert"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "vsbTempSize"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkEnlargeR"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkEnlargeB"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkEnlargeG"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmdCancel"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdClear"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmdBLEND2"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      Begin VB.CommandButton cmdBLEND2 
         Caption         =   "Blend the tow images!!"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   10200
         TabIndex        =   54
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear (fill with the transparent colot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   48
         Top             =   7560
         Width           =   2655
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next step (Overlaying Image)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -65400
         TabIndex        =   46
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton cmdMakeMatrix 
         Caption         =   "Generate matrix code lines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -65280
         TabIndex        =   6
         Top             =   3360
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar pbrMatrixX 
         Height          =   135
         Left            =   -65400
         TabIndex        =   43
         Top             =   3240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   6960
         Width           =   975
      End
      Begin VB.CheckBox chkHighLast 
         Caption         =   "Highlight last character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65400
         TabIndex        =   41
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkBinary 
         Caption         =   "Only binary Matrix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65400
         TabIndex        =   39
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkBoldFont 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65400
         TabIndex        =   38
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkEnlargeG 
         Caption         =   "Enlarge Green channel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6720
         TabIndex        =   37
         Top             =   7080
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkEnlargeB 
         Caption         =   "Enlarge Blue channel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6720
         TabIndex        =   36
         Top             =   7440
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkEnlargeR 
         Caption         =   "Enlarge Red channel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6720
         TabIndex        =   35
         Top             =   6720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.PictureBox picMatColor 
         Height          =   255
         Left            =   -64080
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   34
         Top             =   1920
         Width           =   375
      End
      Begin VB.VScrollBar vsbTempSize 
         Height          =   1335
         Left            =   3480
         Max             =   400
         TabIndex        =   31
         Top             =   6480
         Value           =   100
         Width           =   255
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   7440
         Width           =   2175
      End
      Begin VB.CommandButton cmdLoadPic 
         Caption         =   "Load Picture From File..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   6480
         Width           =   2175
      End
      Begin VB.PictureBox picTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11280
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   28
         Top             =   6600
         Width           =   375
      End
      Begin VB.CommandButton cmdSelTransColor 
         Caption         =   "Select transparemt color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   26
         Top             =   6960
         Width           =   2175
      End
      Begin VB.ComboBox cmbFonts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -64920
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   540
         Width           =   1695
      End
      Begin VB.HScrollBar hsbNumOfLines 
         Enabled         =   0   'False
         Height          =   255
         Left            =   -66240
         Max             =   350
         TabIndex        =   15
         Top             =   7680
         Value           =   13
         Width           =   1815
      End
      Begin VB.CheckBox chkOrg 
         Caption         =   "Organized lines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68280
         TabIndex        =   14
         Top             =   7680
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.HScrollBar hsbDistortion 
         Height          =   255
         Left            =   -71400
         Max             =   200
         TabIndex        =   13
         Top             =   7080
         Width           =   2895
      End
      Begin VB.TextBox txtSize 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -63960
         TabIndex        =   12
         Text            =   "4"
         Top             =   960
         Width           =   375
      End
      Begin VB.HScrollBar hsbLineLength 
         Height          =   255
         Left            =   -71400
         Max             =   200
         TabIndex        =   11
         Top             =   7680
         Value           =   80
         Width           =   2895
      End
      Begin VB.HScrollBar hsbStartingY 
         Height          =   255
         Left            =   -74640
         Max             =   2500
         Min             =   -2500
         TabIndex        =   10
         Top             =   7680
         Value           =   -1500
         Width           =   2895
      End
      Begin VB.HScrollBar hsbHNum 
         Height          =   255
         Left            =   -71400
         Max             =   0
         Min             =   100
         TabIndex        =   9
         Top             =   6600
         Value           =   94
         Width           =   2895
      End
      Begin VB.HScrollBar hsbVSpacing 
         Height          =   255
         Left            =   -74640
         Max             =   500
         TabIndex        =   8
         Top             =   7080
         Value           =   80
         Width           =   2895
      End
      Begin VB.HScrollBar hsbHSpacing 
         Height          =   255
         Left            =   -74640
         Max             =   480
         Min             =   1
         TabIndex        =   7
         Top             =   6600
         Value           =   100
         Width           =   2895
      End
      Begin VB.PictureBox Pic2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   295
         ScaleHeight     =   5835
         ScaleWidth      =   10515
         TabIndex        =   5
         Top             =   360
         Width           =   10575
         Begin VB.Image imgTemp 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DragMode        =   1  'Automatic
            Height          =   1815
            Left            =   3840
            MousePointer    =   15  'Size All
            Top             =   2160
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -74760
         ScaleHeight     =   5835
         ScaleWidth      =   9195
         TabIndex        =   4
         Top             =   360
         Width           =   9255
      End
      Begin MSComctlLib.ProgressBar pbrMatrixY 
         Height          =   510
         Left            =   -65400
         TabIndex        =   44
         Top             =   3345
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   900
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "If you get a sucky blended image play with this checkboxes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   53
         Top             =   6240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":0038
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -65400
         TabIndex        =   52
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   3720
         TabIndex        =   51
         Top             =   6480
         Width           =   210
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "400%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   3720
         TabIndex        =   50
         Top             =   7560
         Width           =   735
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<---RESIZE IMAGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   3840
         TabIndex        =   49
         Top             =   6960
         Width           =   1365
      End
      Begin VB.Image imgDefault 
         Height          =   135
         Left            =   120
         Picture         =   "Form1.frx":010A
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         Caption         =   "Base Matrix color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   -65400
         TabIndex        =   33
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label lbldesc 
         Caption         =   $"Form1.frx":460C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Index           =   10
         Left            =   10320
         TabIndex        =   32
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   1560
         X2              =   1440
         Y1              =   7200
         Y2              =   7440
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   1440
         X2              =   1320
         Y1              =   7440
         Y2              =   7200
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   1440
         X2              =   1440
         Y1              =   6840
         Y2              =   7440
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current transparent color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   9360
         TabIndex        =   27
         Top             =   6600
         Width           =   1785
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         Caption         =   "Amount of columns:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   -66240
         TabIndex        =   25
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         Caption         =   "Font:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   -65400
         TabIndex        =   23
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lbldesc 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   -64440
         TabIndex        =   22
         Top             =   1005
         Width           =   345
      End
      Begin VB.Label lbldesc 
         Caption         =   "Length of cloumn:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -71400
         TabIndex        =   21
         Top             =   7440
         Width           =   2535
      End
      Begin VB.Label lbldesc 
         Caption         =   "Distortion (both sides):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -71400
         TabIndex        =   20
         Top             =   6840
         Width           =   2535
      End
      Begin VB.Label lbldesc 
         Caption         =   "Chance of highlighting a character:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -71400
         TabIndex        =   19
         Top             =   6360
         Width           =   2535
      End
      Begin VB.Label lbldesc 
         Caption         =   "Minimal Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   18
         Top             =   7440
         Width           =   2535
      End
      Begin VB.Label lbldesc 
         Caption         =   "Vertical Spacing:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   17
         Top             =   6840
         Width           =   2535
      End
      Begin VB.Label lbldesc 
         Caption         =   "Spacing between columns:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   16
         Top             =   6360
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Blend the two images!!"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pbrBlendX 
      Height          =   3015
      Left            =   12480
      TabIndex        =   45
      Top             =   5040
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   5318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label lbldesc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "It's worth waiting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   12720
      TabIndex        =   47
      Top             =   6000
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12360
      TabIndex        =   1
      Top             =   4320
      Width           =   150
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P As Boolean
Dim TransSelect As Boolean
Dim UserStop As Boolean
Dim TransColor

Private Sub chkOrg_Click()
    hsbNumOfLines.Enabled = Not hsbNumOfLines.Enabled
    hsbHSpacing.Enabled = Not hsbHSpacing.Enabled
End Sub

Private Sub cmbFonts_Click()
    If cmbFonts <> "" Then cmbFonts.Font.Name = cmbFonts.Text
End Sub

Private Sub cmdBLEND2_Click()
    Command1_Click
End Sub

Private Sub cmdCancel_Click()
    imgTemp.Visible = False
    imgTemp.Stretch = False
    vsbTempSize.Enabled = False
    vsbTempSize.Value = 100
    imgTemp.Picture = LoadPicture("")
End Sub

Private Sub cmdClear_Click()
    Pic2.BackColor = TransColor
    Pic2.Cls
End Sub

Private Sub cmdInsert_Click()
    If imgTemp.Picture = 0 Then cmdLoadPic_Click: Exit Sub
    Pic2.PaintPicture imgTemp.Picture, imgTemp.Left, imgTemp.Top, imgTemp.Width, imgTemp.Height
    imgTemp.Visible = False
    imgTemp.Stretch = False
    vsbTempSize.Enabled = False
    vsbTempSize.Value = 100
    imgTemp.Picture = LoadPicture("")
End Sub

Private Sub cmdLoadPic_Click()
    CDialog.Filter = "Pictures (*.bmp;*.jpg)|*.bmp;*.jpg"
    CDialog.ShowOpen
    If CDialog.FileName <> "" Then
        imgTemp.Picture = LoadPicture(CDialog.FileName)
        imgTemp.Visible = True
        imgTemp.Stretch = True
        vsbTempSize.Enabled = True
    End If
End Sub

Private Sub cmdMakeMatrix_Click()
On Error GoTo FontERR
Dim I, K, StartY, C, Ch, writeX, LineNum
Dim Base As RGBColor
    Pic1.Cls
    Pic1.Font.Size = Val(txtSize)
    Pic1.FontBold = chkBoldFont.Value
    Pic1.FontName = "fixedsys"
    Pic1.FontItalic = False
    Pic1.Font.Underline = False
    
    Base = GetColorFromLong(picMatColor.BackColor)
    
    If cmbFonts <> "" Then Pic1.FontName = cmbFonts
    If chkOrg.Value = 1 Then LineNum = Pic1.ScaleWidth / (hsbHSpacing.Value)
    If chkOrg.Value = 0 Then LineNum = hsbNumOfLines.Value
    K = -1
    Do While K < LineNum
        K = K + 1
        Randomize
        StartY = hsbStartingY.Value + Rnd * 2000
        If chkOrg.Value = 0 Then writeX = Rnd * Pic1.ScaleWidth
        
        pbrMatrixX.Max = LineNum + 1
        pbrMatrixX.Value = K
        
        For I = 0 To hsbLineLength
            pbrMatrixY.Max = hsbLineLength
            pbrMatrixY.Value = I
            With Pic1
                .CurrentY = StartY + I * (hsbVSpacing.Value)
                If chkOrg.Value = 1 Then
                    .CurrentX = K * (hsbHSpacing.Value) - hsbDistortion.Value + Rnd * (hsbDistortion * 2)
                Else: .CurrentX = writeX - hsbDistortion.Value + Rnd * (hsbDistortion * 2)
                End If
                'uncomment the next 2 lines and you'll have a nice effect
'                .CurrentX = (I) * (hsbHSpacing.Value)
'                .CurrentY = (K + K * I / 30) * (hsbHSpacing.Value)
                Randomize
                If I > 12 Then
                    If Rnd * 100 > hsbHNum Then .ForeColor = RGB(Base.R + 70, Base.G + 70, Base.B + 70) 'RGB(50, 175 + Rnd * 50, 25)
                Else
                    .ForeColor = RGB(I * Base.R / 12, I * Base.G / 12, I * Base.B / 12)
                End If
TryAgain:
                Randomize
                If chkBinary.Value = 0 Then
                    Ch = Chr(Rnd * 42 + 48)
                    If Ch = ":" Or Ch = "_" Or Ch = ";" Or Ch = "=" Or Ch = "-" Or Ch = "+" Or Ch = "`" Then GoTo TryAgain
                Else: Randomize: Ch = Trim(Str(Round(Rnd * 1)))
                End If
                Pic1.Print Ch
                
                .ForeColor = picMatColor.BackColor
            End With
        Next
        
        If chkHighLast.Value = 1 Then
        Pic1.CurrentY = StartY + I * (hsbVSpacing.Value)
            If chkOrg.Value = 1 Then
                Pic1.CurrentX = K * (hsbHSpacing.Value) - hsbDistortion.Value + Rnd * (hsbDistortion.Value * 2)
            Else: Pic1.CurrentX = writeX - hsbDistortion.Value + Rnd * (hsbDistortion * 2)
            End If
            C = 150 + Rnd * 105
            Pic1.ForeColor = RGB(Base.R + 140, Base.G + 140, Base.B + 140)
TryAgain2:
            Randomize
            If chkBinary.Value = 0 Then
                Ch = Chr(Rnd * 42 + 48)
                If Ch = ":" Or Ch = "_" Or Ch = ";" Or Ch = "=" Or Ch = "-" Or Ch = "+" Or Ch = "`" Then GoTo TryAgain2
            Else: Ch = Trim(Str(Round(Rnd * 1)))
            End If
            Pic1.Print Ch
        End If
        Pic1.ForeColor = RGB(Base.R, Base.G, Base.B)
    Loop
    Exit Sub
FontERR:     MsgBox "You don't have this font. If its ""IrisUPC"" - copy it from the .zip into the font directory"
End Sub

Private Sub cmdNext_Click()
    SSTab1.Tab = 1
End Sub

Private Sub cmdSave_Click()
    CDialog.Filter = "Bitmap|*.bmp"
    CDialog.ShowSave
    SavePicture Pic1.Image, CDialog.FileName & ".bmp"
End Sub

Private Sub cmdSelTransColor_Click()
    If imgTemp.Picture = 0 Then
        TransSelect = True
        Pic2.MousePointer = 2
        imgTemp.DragMode = 0
    Else: MsgBox "There is an un-inserted picture.." & vbNewLine & "click the ""Insert picture"" button to insert it."
    End If
End Sub

Private Sub Command1_Click()
    Dim X, Y
    Command1.Enabled = False
    Command3.Enabled = True
    
    pbrBlendX.Max = Pic1.ScaleWidth - 16
    For X = 0 To Pic2.ScaleWidth - 16 Step 15
        pbrBlendX.Value = X
        DoEvents
        For Y = 0 To Pic2.ScaleHeight - 16 Step 15
            If Pic2.Point(X, Y) <> TransColor Then
                If Pic1.Point(X, Y) > 0 And Pic1.Point(X, Y) <> Pic1.BackColor Then
                    Color1 = GetColorFromLong(Pic1.Point(X, Y))
                    Me.BackColor = RGB(Color1.R, Color1.G, Color1.B)
                    Color2 = GetColorFromLong(Pic2.Point(X, Y))
                    Me.BackColor = Pic2.Point(X, Y)
                    NewColor = Blend(Color1, Color2, chkEnlargeR.Value + 1, chkEnlargeG.Value + 1, chkEnlargeB.Value + 1)
                    Pic1.ForeColor = RGB(NewColor.R, NewColor.G, NewColor.B)
                    Me.BackColor = Pic1.ForeColor
                    Pic1.PSet (X, Y)
                End If
            End If
            Label1 = "X: " & X & vbNewLine & "Y: " & Y
        Next
        If UserStop = True Then Exit For
    Next
    Me.BackColor = vbButtonFace
    UserStop = False
    Command1.Enabled = True
End Sub

Private Sub Command3_Click()
    UserStop = True
    Command3.Enabled = False
End Sub

'original ForeColor=&H00008000& = rgb(128,128,0)
Private Sub Form_Load()
    Dim A
    Show
    For A = 0 To Screen.FontCount
        cmbFonts.AddItem Screen.Fonts(A)
    Next
    cmbFonts.Text = "IrisUPC"
    Pic2.Move Pic2.Left, Pic2.Top, Pic1.Width, Pic1.Height
    picTrans.BackColor = Pic2.BackColor
    TransColor = Pic2.BackColor
    picMatColor.BackColor = RGB(0, 128, 0)
    Pic2.PaintPicture imgDefault.Picture, 600, 400, 7700, 5500
    CDialog.InitDir = App.Path
    SSTab1.Tab = 0
    cmdMakeMatrix_Click
End Sub

Private Sub imgTemp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgTemp.Tag = "D" Then imgTemp.Move X, Y
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox "HEX color: #" & Hex$(Pic1.Point(X, Y)) '& vbNewLine & RGB
End Sub

Private Sub Pic2_DragDrop(Source As Control, X As Single, Y As Single)
    imgTemp.Move X, Y
End Sub

Private Sub Pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 4 And TransSelect = False Then
        P = True
        Pic2.ForeColor = vbWhite
        If Button = 2 Then Pic2.ForeColor = TransColor
        Pic2_MouseMove Button, Shift, X, Y
    End If
    If TransSelect = True Then
        TransColor = Pic2.Point(X, Y)
        Pic2.MousePointer = 1
        TransSelect = False
        picTrans.BackColor = TransColor
        imgTemp.DragMode = 1
    End If
    If Button = 4 Then
        Pic2.BackColor = TransColor
        Pic2.Cls
    End If
End Sub

Private Sub Pic2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If P = True Then
        Pic2.DrawWidth = 15
        Pic2.PSet (X, Y)
    End If
    Print X, Y
End Sub

Private Sub Pic2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    P = False
End Sub

Private Sub picMatColor_Click()
    CDialog.Color = picMatColor.BackColor
    CDialog.DialogTitle = "Select Matrix color"
    CDialog.Flags = cdlCCRGBInit
    CDialog.ShowColor
    picMatColor.BackColor = CDialog.Color
End Sub


Private Sub tmrTitle_Timer()
    Me.Caption = "Matrix Blender by Michael Vainshtein" & String(Int(Rnd * 2), "_")
End Sub

Private Sub vsbTempSize_Change()
    imgTemp.Width = vsbTempSize.Value * (imgTemp.Picture.Width / 100)
    imgTemp.Height = vsbTempSize.Value * (imgTemp.Picture.Height / 100)
End Sub
