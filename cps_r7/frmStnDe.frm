VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStnDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Station Detail Screen"
   ClientHeight    =   11835
   ClientLeft      =   195
   ClientTop       =   900
   ClientWidth     =   15330
   FillColor       =   &H00FFFFC0&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFC0&
   Icon            =   "frmStnDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11835
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin Threed.SSPanel pnlLoadRate 
      Height          =   810
      Left            =   9600
      TabIndex        =   75
      Top             =   8040
      Width           =   5040
      _Version        =   65536
      _ExtentX        =   8890
      _ExtentY        =   1429
      _StockProps     =   15
      Caption         =   "Load Rate SetPoint"
      ForeColor       =   4210816
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   8
      Begin Threed.SSPanel pnlLoadRateSp 
         Height          =   375
         Left            =   1920
         TabIndex        =   77
         ToolTipText     =   "Scale Display"
         Top             =   90
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   661
         _StockProps     =   15
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Begin VB.TextBox txtLoadRateSp 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   60
            TabIndex        =   78
            TabStop         =   0   'False
            Text            =   "40.0"
            ToolTipText     =   "New Load Rate"
            Top             =   60
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmdLoadRateUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4050
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmStnDe.frx":57E2
         TabIndex        =   76
         ToolTipText     =   "Update MFC SetPoints"
         Top             =   157
         UseMaskColor    =   -1  'True
         Width           =   870
      End
   End
   Begin MSComctlLib.ImageList SmallImagesDisabled 
      Left            =   7080
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":5B24
            Key             =   "bargraph"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":5E76
            Key             =   "xygraph"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImagesHot 
      Left            =   6720
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":61C8
            Key             =   "bargraph"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":651A
            Key             =   "xygraph"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImagesNormal 
      Left            =   6360
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":686C
            Key             =   "bargraph"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":6BBE
            Key             =   "xygraph"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrStnDetail 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   60
      Top             =   3240
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1588
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      ImageList       =   "imgStnDetailNormal"
      DisabledImageList=   "imgStnDetailDisabled"
      HotImageList    =   "imgStnDetailHot"
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin Threed.SSPanel pnlXYGraphs 
      Height          =   5670
      Left            =   0
      TabIndex        =   80
      Top             =   4200
      Width           =   11985
      _Version        =   65536
      _ExtentX        =   21140
      _ExtentY        =   10001
      _StockProps     =   15
      ForeColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin MSChart20Lib.MSChart chtStnChart 
         Height          =   5505
         Left            =   90
         OleObjectBlob   =   "frmStnDe.frx":6F10
         TabIndex        =   81
         Top             =   120
         Visible         =   0   'False
         Width           =   6765
      End
   End
   Begin VB.Timer tmrXYGraphs 
      Interval        =   1000
      Left            =   9120
      Top             =   9360
   End
   Begin MSComctlLib.ImageList imgStnDetailDisabled 
      Left            =   12840
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":9B5F
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":B6B1
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":D203
            Key             =   "alarmlog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":ED55
            Key             =   "ootlog"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":108A7
            Key             =   "statsum"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":123F9
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":1434B
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":15E9D
            Key             =   "continue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":179EF
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":19541
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":1B093
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":1CBE5
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":1E737
            Key             =   "fuelsupply"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":20289
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":21DDB
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":2392D
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":2547F
            Key             =   "opercomment"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgStnDetailHot 
      Left            =   13440
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":26FD1
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":28B23
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":2A675
            Key             =   "alarmlog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":2C1C7
            Key             =   "ootlog"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":2DD19
            Key             =   "statsum"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":2F86B
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":317BD
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":3330F
            Key             =   "continue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":34E61
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":369B3
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":38505
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":3A057
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":3BBA9
            Key             =   "fuelsupply"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":3D6FB
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":3F24D
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":40D9F
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":428F1
            Key             =   "opercomment"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgStnDetailNormal 
      Left            =   12360
      Top             =   9840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":44443
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":45F95
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":47AE7
            Key             =   "alarmlog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":49639
            Key             =   "ootlog"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":4B18B
            Key             =   "statsum"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":4CCDD
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":4EC2F
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":50781
            Key             =   "continue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":522D3
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":53E25
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":55977
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":574C9
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":5901B
            Key             =   "fuelsupply"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":5AB6D
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":5C6BF
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":5E211
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStnDe.frx":5FD63
            Key             =   "opercomment"
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel pbxRptName 
      Height          =   2220
      Left            =   15000
      TabIndex        =   62
      Top             =   8880
      Width           =   5040
      _Version        =   65536
      _ExtentX        =   8890
      _ExtentY        =   3916
      _StockProps     =   15
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtRptMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   73
         Text            =   "frmStnDe.frx":618B5
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtRptName2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   65
         ToolTipText     =   "25 characters max."
         Top             =   645
         Width           =   3000
      End
      Begin VB.TextBox txtRptName3 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   64
         ToolTipText     =   "25 characters max."
         Top             =   930
         Width           =   3000
      End
      Begin VB.TextBox txtRptName1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   63
         ToolTipText     =   "25 characters max."
         Top             =   360
         Width           =   3000
      End
      Begin Threed.SSCommand cmdApproved 
         Height          =   735
         Left            =   4080
         TabIndex        =   66
         Tag             =   "Approved"
         ToolTipText     =   "Report File Name Text is Valid"
         Top             =   405
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BevelWidth      =   0
         Outline         =   0   'False
         Picture         =   "frmStnDe.frx":618BF
      End
      Begin Threed.SSCommand cmdValidate 
         Height          =   735
         Left            =   3240
         TabIndex        =   67
         Tag             =   "Validate"
         ToolTipText     =   "Validate Report File Name Text"
         Top             =   405
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Outline         =   0   'False
         Picture         =   "frmStnDe.frx":670B1
      End
      Begin Threed.SSCommand cmdApproved_OK 
         Height          =   735
         Left            =   3720
         TabIndex        =   161
         ToolTipText     =   "Report File Name Text is Valid"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Outline         =   0   'False
         Picture         =   "frmStnDe.frx":6C8A3
      End
      Begin Threed.SSCommand cmdApproved_No 
         Height          =   735
         Left            =   2880
         TabIndex        =   162
         ToolTipText     =   "Report File Name Text is NOT Valid"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Outline         =   0   'False
         Picture         =   "frmStnDe.frx":72095
      End
      Begin VB.Label Label6 
         Caption         =   "Validate File Name"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   69
         Top             =   120
         Width           =   1680
      End
      Begin VB.Label lblRptNameOper 
         Caption         =   "Enter Report File Name"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   2880
      End
   End
   Begin Threed.SSPanel pnlStnSeq 
      Height          =   900
      Left            =   10080
      TabIndex        =   30
      ToolTipText     =   "Active Sequence Description"
      Top             =   7200
      Width           =   5040
      _Version        =   65536
      _ExtentX        =   8881
      _ExtentY        =   1587
      _StockProps     =   15
      Caption         =   "Station Sequence ## is Active "
      ForeColor       =   16512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   8
      Begin VB.TextBox txtSeqMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   360
         Left            =   120
         TabIndex        =   31
         Text            =   "Station Sequence Message Window"
         ToolTipText     =   "Active Step Description"
         Top             =   120
         Width           =   4800
      End
   End
   Begin Threed.SSPanel pnlCanVentOvr 
      Height          =   585
      Left            =   14400
      TabIndex        =   29
      ToolTipText     =   "CANVENT Flow Switch Override Status"
      Top             =   5760
      Width           =   5040
      _Version        =   65536
      _ExtentX        =   8890
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "CanVent Override is Active (12345 of 29999 sec)"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
   End
   Begin VB.PictureBox pbxBottom 
      Align           =   2  'Align Bottom
      Height          =   460
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   15270
      TabIndex        =   45
      Top             =   11370
      Width           =   15330
      Begin Threed.SSPanel pnlAlarms 
         Height          =   405
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   5190
         _Version        =   65536
         _ExtentX        =   9155
         _ExtentY        =   714
         _StockProps     =   15
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Begin Threed.SSPanel pnlEstop 
            Height          =   255
            Left            =   80
            TabIndex        =   47
            ToolTipText     =   "EMERGENCY Stop Pressed"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "ESTOP"
            ForeColor       =   -2147483630
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlFlow 
            Height          =   255
            Left            =   915
            TabIndex        =   48
            ToolTipText     =   "Loss of Flow"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "FLOW"
            ForeColor       =   -2147483630
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlBtn20 
            Height          =   255
            Left            =   2595
            TabIndex        =   49
            ToolTipText     =   "20% Butane LEL Alarm"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "LEL 20"
            ForeColor       =   -2147483630
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlDoors 
            Height          =   255
            Left            =   1755
            TabIndex        =   50
            ToolTipText     =   "Loss of Vacuum"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "DOORS"
            ForeColor       =   -2147483630
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlComms 
            Height          =   255
            Left            =   3435
            TabIndex        =   51
            ToolTipText     =   "Communication Error"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "COMMS"
            ForeColor       =   -2147483630
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlMix 
            Height          =   255
            Left            =   4270
            TabIndex        =   52
            ToolTipText     =   "Communication Error"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "MIX"
            ForeColor       =   -2147483630
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
            BevelOuter      =   1
         End
      End
      Begin Threed.SSPanel pnlPurgeAir 
         Height          =   405
         Left            =   9360
         TabIndex        =   53
         Top             =   0
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "purge air"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodShowPct    =   0   'False
      End
      Begin Threed.SSPanel pnlMessageFrame 
         Height          =   405
         Left            =   5190
         TabIndex        =   54
         Top             =   0
         Width           =   4200
         _Version        =   65536
         _ExtentX        =   7408
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "message"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodShowPct    =   0   'False
         Begin Threed.SSPanel pnlMessage 
            Height          =   280
            Left            =   60
            TabIndex        =   72
            Top             =   60
            Width           =   4080
            _Version        =   65536
            _ExtentX        =   7197
            _ExtentY        =   494
            _StockProps     =   15
            Caption         =   "message"
            ForeColor       =   -2147483630
            BackColor       =   8454143
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   0
            BorderWidth     =   0
            BevelOuter      =   0
         End
      End
   End
   Begin VB.PictureBox pbxVertical 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7230
      Left            =   12405
      ScaleHeight     =   7230
      ScaleWidth      =   2925
      TabIndex        =   33
      Top             =   4140
      Width           =   2925
      Begin Threed.SSPanel pnlProfile 
         Height          =   1185
         Left            =   0
         TabIndex        =   84
         ToolTipText     =   "Current Time Remaining"
         Top             =   840
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   2090
         _StockProps     =   15
         Caption         =   "Purge Profile"
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel pnlStepMax 
            Height          =   375
            Left            =   1320
            TabIndex        =   85
            Top             =   90
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "888.8"
            ForeColor       =   12583104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlStepActual 
            Height          =   375
            Left            =   2070
            TabIndex        =   86
            Top             =   90
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "888.8"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlMinutesMax 
            Height          =   375
            Left            =   1320
            TabIndex        =   87
            Top             =   465
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   12583104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlMinutesActual 
            Height          =   375
            Left            =   2070
            TabIndex        =   88
            Top             =   465
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin VB.Label lblStepMinutes 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "max/minutes"
            Height          =   270
            Left            =   0
            TabIndex        =   90
            Top             =   525
            Width           =   1395
         End
         Begin VB.Label lblProfileStep 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "max/step"
            Height          =   270
            Left            =   0
            TabIndex        =   89
            Top             =   165
            Width           =   1395
         End
      End
      Begin Threed.SSPanel pnlWaterBath 
         Height          =   1185
         Left            =   0
         TabIndex        =   154
         ToolTipText     =   "Current Time Remaining"
         Top             =   600
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   2090
         _StockProps     =   15
         Caption         =   "WaterBath Temperatures"
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel pnlSuperSp 
            Height          =   375
            Left            =   1320
            TabIndex        =   155
            Top             =   90
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "888.8"
            ForeColor       =   12583104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlWbSp 
            Height          =   375
            Left            =   1320
            TabIndex        =   156
            Top             =   465
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   12583104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlWbPv 
            Height          =   375
            Left            =   2070
            TabIndex        =   157
            Top             =   465
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlSuperPv 
            Height          =   375
            Left            =   2070
            TabIndex        =   158
            Top             =   90
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin VB.Label lblSpPv 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "WB SP/PV"
            Height          =   270
            Left            =   0
            TabIndex        =   160
            Top             =   525
            Width           =   1395
         End
         Begin VB.Label lblSuperSP 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PID SP/PV"
            Height          =   270
            Left            =   0
            TabIndex        =   159
            Top             =   165
            Width           =   1395
         End
      End
      Begin MSComctlLib.TabStrip tabsWhichGraph 
         Height          =   375
         Left            =   360
         TabIndex        =   82
         Top             =   5280
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   661
         TabWidthStyle   =   2
         Style           =   1
         TabFixedWidth   =   2117
         ImageList       =   "SmallImagesNormal"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Net Grams"
               Key             =   "netgrams"
               Object.ToolTipText     =   "Show XY Graph of Net Grams of Butane"
               ImageVarType    =   8
               ImageKey        =   "xygraph"
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "MFC Flow"
               Key             =   "mfcflow"
               Object.ToolTipText     =   "Show Bar Graphs and Numerical Values of MFC Flows"
               ImageVarType    =   8
               ImageKey        =   "bargraph"
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TabStrip tabsLegend 
         Height          =   375
         Left            =   0
         TabIndex        =   83
         Top             =   5000
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   661
         TabWidthStyle   =   2
         Style           =   1
         TabFixedWidth   =   2117
         ImageList       =   "SmallImagesNormal"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Legend"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "No Legend"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel pnlLcPressBox 
         Height          =   375
         Left            =   0
         TabIndex        =   146
         ToolTipText     =   "LeakCheck Pressure Reading"
         Top             =   2160
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   661
         _StockProps     =   15
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel SSPanel2 
            Height          =   345
            Left            =   -1920
            TabIndex        =   147
            Top             =   120
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "1"
            ForeColor       =   0
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlLcPress 
            Height          =   285
            Left            =   1120
            TabIndex        =   148
            ToolTipText     =   "Current LeakCheck Pressure Reading"
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "012345"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
         Begin VB.Label lblLcPress 
            BackStyle       =   0  'Transparent
            Caption         =   "LC Press"
            Height          =   285
            Left            =   120
            TabIndex        =   150
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label lblLcPressUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "psi"
            Height          =   285
            Left            =   2200
            TabIndex        =   149
            Top             =   90
            Width           =   1035
         End
      End
      Begin Threed.SSPanel pnlLeakcheck 
         Height          =   1215
         Left            =   0
         TabIndex        =   118
         ToolTipText     =   "Leakcheck Information"
         Top             =   4440
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   2143
         _StockProps     =   15
         Caption         =   "Leakcheck"
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel pnlLcStepDesc 
            Height          =   390
            Index           =   0
            Left            =   120
            TabIndex        =   119
            Top             =   90
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   679
            _StockProps     =   15
            Caption         =   "Pressurizing"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlLcStepDesc 
            Height          =   390
            Index           =   1
            Left            =   120
            TabIndex        =   120
            Top             =   480
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   679
            _StockProps     =   15
            Caption         =   "Pressurizing"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
      End
      Begin Threed.SSPanel pnlCourse 
         Height          =   375
         Left            =   0
         TabIndex        =   117
         Top             =   840
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Course 1 of 2"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlDelay 
         Height          =   1185
         Left            =   0
         TabIndex        =   35
         ToolTipText     =   "Current Time Remaining"
         Top             =   240
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   2090
         _StockProps     =   15
         Caption         =   "Delay"
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel pnlToGo 
            Height          =   375
            Left            =   1400
            TabIndex        =   36
            ToolTipText     =   "Delay Remaining in Seconds"
            Top             =   90
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "012345"
            ForeColor       =   -2147483646
            BackColor       =   -2147483626
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlTotal 
            Height          =   375
            Left            =   1395
            TabIndex        =   37
            ToolTipText     =   "Total Delay in Seconds"
            Top             =   465
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "012345"
            ForeColor       =   -2147483646
            BackColor       =   -2147483626
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   270
            Left            =   30
            TabIndex        =   39
            Top             =   545
            Width           =   1400
         End
         Begin VB.Label lblToGo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "To Go"
            Height          =   270
            Left            =   30
            TabIndex        =   38
            Top             =   170
            Width           =   1395
         End
      End
      Begin Threed.SSPanel pnlScale 
         Height          =   1185
         Left            =   0
         TabIndex        =   40
         Top             =   2520
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   2090
         _StockProps     =   15
         ForeColor       =   4210816
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel pnlWtAux 
            Height          =   375
            Left            =   855
            TabIndex        =   41
            ToolTipText     =   "Current Aux. Scale Reading"
            Top             =   465
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "00,000.00"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlWtPri 
            Height          =   375
            Left            =   855
            TabIndex        =   42
            ToolTipText     =   "Current Primary Scale Reading"
            Top             =   90
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "00,000.00"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlChgAux 
            Height          =   375
            Left            =   1980
            TabIndex        =   111
            ToolTipText     =   "Aux  scale weight change during current Load or Purge"
            Top             =   465
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "0,000.00"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlChgPri 
            Height          =   375
            Left            =   1980
            TabIndex        =   112
            ToolTipText     =   "Primary scale weight change during current Load or Purge"
            Top             =   90
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "0,000.00"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin VB.Label lblScaleDesc 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Scale"
            ForeColor       =   &H00004080&
            Height          =   270
            Left            =   120
            TabIndex        =   115
            Top             =   890
            Width           =   735
         End
         Begin VB.Label lblScaleWeight 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
            ForeColor       =   &H00004080&
            Height          =   270
            Left            =   855
            TabIndex        =   114
            Top             =   890
            Width           =   1125
         End
         Begin VB.Label lblScaleChange 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Change"
            ForeColor       =   &H00004080&
            Height          =   270
            Left            =   1980
            TabIndex        =   113
            Top             =   890
            Width           =   885
         End
         Begin VB.Label lblPriScale 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pri. #"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   79
            Top             =   150
            Width           =   735
         End
         Begin VB.Label lblAuxScale 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aux. #8"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   525
            Width           =   735
         End
      End
      Begin Threed.SSPanel pnlCycle 
         Height          =   375
         Left            =   0
         TabIndex        =   44
         Top             =   600
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Cycle 1 of 2"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlTestTime 
         Height          =   375
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "Elapsed Test Time in HH:MM:SS.ss"
         Top             =   120
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Elapsed :  DD days HH:MM:SS"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSCommand cmdUseTC 
         Height          =   435
         Left            =   2100
         TabIndex        =   70
         ToolTipText     =   "Toggle Thermocouple Display ON / off"
         Top             =   5680
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "OFF"
         ForeColor       =   -2147483646
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdThermo 
         Height          =   435
         Left            =   0
         TabIndex        =   71
         ToolTipText     =   "Display Common Thermocouples"
         Top             =   5280
         Width           =   2880
         _Version        =   65536
         _ExtentX        =   5080
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Common TC's"
         ForeColor       =   -2147483646
         Outline         =   0   'False
         Picture         =   "frmStnDe.frx":77887
      End
      Begin Threed.SSPanel pnlJobDuration 
         Height          =   375
         Left            =   0
         TabIndex        =   116
         ToolTipText     =   "Estimated Job Duration in Days Hours & Minutes"
         Top             =   360
         Width           =   5040
         _Version        =   65536
         _ExtentX        =   8890
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Estimated  Job Duration in Days Hours Minutes"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlPurgeDpBox 
         Height          =   375
         Left            =   0
         TabIndex        =   121
         ToolTipText     =   "Purge Differential Pressure Reading"
         Top             =   2040
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   661
         _StockProps     =   15
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel SSPanel12 
            Height          =   345
            Left            =   -1920
            TabIndex        =   122
            Top             =   120
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "1"
            ForeColor       =   0
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlPurgeDp 
            Height          =   285
            Left            =   1120
            TabIndex        =   123
            ToolTipText     =   "Current Purge Differential Pressure Reading"
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "012345"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
         Begin VB.Label lblPurgeDpUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "in H2O"
            Height          =   285
            Left            =   2200
            TabIndex        =   125
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label lblPurgeDp 
            BackStyle       =   0  'Transparent
            Caption         =   "Purge DP"
            Height          =   285
            Left            =   120
            TabIndex        =   124
            Top             =   90
            Width           =   1035
         End
      End
      Begin Threed.SSPanel pnlTC 
         Height          =   495
         Left            =   0
         TabIndex        =   126
         ToolTipText     =   "Thermocouple readings"
         Top             =   6360
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   873
         _StockProps     =   15
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel pnlTC1 
            Height          =   345
            Left            =   345
            TabIndex        =   127
            ToolTipText     =   "Thermocouple One Value"
            Top             =   90
            Width           =   720
            _Version        =   65536
            _ExtentX        =   1270
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "199.9"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   1
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlTC2 
            Height          =   345
            Left            =   1320
            TabIndex        =   128
            ToolTipText     =   "Thermocouple Two Value"
            Top             =   90
            Width           =   720
            _Version        =   65536
            _ExtentX        =   1270
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "97.4"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   1
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnlLblTC1 
            Height          =   345
            Left            =   90
            TabIndex        =   129
            Top             =   90
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "1"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlLblTC2 
            Height          =   345
            Left            =   1080
            TabIndex        =   130
            Top             =   90
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "2"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
         End
      End
   End
   Begin VB.Timer tmrScreen 
      Interval        =   500
      Left            =   6615
      Top             =   7920
   End
   Begin Threed.SSPanel pnlBarGraphs 
      Height          =   5700
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   5865
      _Version        =   65536
      _ExtentX        =   10345
      _ExtentY        =   10054
      _StockProps     =   15
      ForeColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin Threed.SSPanel pnlNitrogen 
         Height          =   5550
         Left            =   2940
         TabIndex        =   1
         ToolTipText     =   "Set Point / Process Variables "
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   9798
         _StockProps     =   15
         Caption         =   "Nitrogen"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Alignment       =   8
         Begin VB.TextBox txtNitCV 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Text            =   "000.0"
            ToolTipText     =   "Actual Nitrogen SP Values"
            Top             =   4900
            Width           =   600
         End
         Begin VB.TextBox txtNitPV 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   750
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "00.00"
            ToolTipText     =   "Actual Nitrogen PV Values"
            Top             =   4900
            Width           =   600
         End
         Begin Threed.SSPanel pnlNitCV 
            Height          =   4395
            Left            =   240
            TabIndex        =   4
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodColor      =   49152
            FloodShowPct    =   0   'False
         End
         Begin Threed.SSPanel pnlNitPV 
            Height          =   4395
            Left            =   850
            TabIndex        =   5
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodShowPct    =   0   'False
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   255
            Left            =   260
            TabIndex        =   6
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "SP"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   255
            Left            =   870
            TabIndex        =   7
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "PV"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel pnlButane 
         Height          =   5550
         Left            =   4365
         TabIndex        =   8
         ToolTipText     =   "Set Point / Process Variables "
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   9798
         _StockProps     =   15
         Caption         =   "Butane"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Alignment       =   8
         Begin VB.TextBox txtBtnPV 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   750
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "199.9"
            ToolTipText     =   "Actual Butane PV Values"
            Top             =   4900
            Width           =   600
         End
         Begin VB.TextBox txtBtnCV 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "199.9"
            ToolTipText     =   "Actual Butane CV Values"
            Top             =   4900
            Width           =   600
         End
         Begin Threed.SSPanel pnlBtnCV 
            Height          =   4395
            Left            =   240
            TabIndex        =   11
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodColor      =   49152
            FloodShowPct    =   0   'False
         End
         Begin Threed.SSPanel pnlBtnPV 
            Height          =   4395
            Left            =   850
            TabIndex        =   12
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodShowPct    =   0   'False
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   255
            Left            =   260
            TabIndex        =   13
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "SP"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   255
            Left            =   870
            TabIndex        =   14
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "PV"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel pnlPurge 
         Height          =   5550
         Left            =   1515
         TabIndex        =   15
         ToolTipText     =   "Set Point / Process Variables "
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   9798
         _StockProps     =   15
         Caption         =   "Purge"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Alignment       =   8
         Begin Threed.SSPanel pnlPurPV 
            Height          =   4395
            Left            =   850
            TabIndex        =   19
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodShowPct    =   0   'False
         End
         Begin VB.TextBox txtPurPV 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   750
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "888.88"
            ToolTipText     =   "Actual Purge PV Values"
            Top             =   4900
            Width           =   600
         End
         Begin VB.TextBox txtPurCV 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "999.99"
            ToolTipText     =   "Actual Purge SP Values"
            Top             =   4900
            Width           =   600
         End
         Begin Threed.SSPanel pnlPurCV 
            Height          =   4395
            Left            =   240
            TabIndex        =   18
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodColor      =   49152
            FloodShowPct    =   0   'False
         End
         Begin Threed.SSPanel SSPanel15 
            Height          =   255
            Left            =   260
            TabIndex        =   20
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "SP"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   870
            TabIndex        =   21
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "PV"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel pnlTargetActual 
         Height          =   5555
         Left            =   90
         TabIndex        =   22
         ToolTipText     =   "Target / Actual Reading"
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   9798
         _StockProps     =   15
         Caption         =   " Target  Actual"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Alignment       =   2
         Begin VB.TextBox txtActual 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   750
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "100.0"
            ToolTipText     =   "Actual Value"
            Top             =   4900
            Width           =   600
         End
         Begin VB.TextBox txtTarget 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "100.0"
            ToolTipText     =   "Actual Target Value"
            Top             =   4900
            Width           =   600
         End
         Begin Threed.SSPanel pnlTarget 
            Height          =   4395
            Left            =   240
            TabIndex        =   23
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   65280
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodColor      =   49152
            FloodShowPct    =   0   'False
         End
         Begin Threed.SSPanel pnlActual 
            Height          =   4395
            Left            =   850
            TabIndex        =   24
            Top             =   420
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   7752
            _StockProps     =   15
            ForeColor       =   65280
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelInner      =   1
            FloodType       =   4
            FloodShowPct    =   0   'False
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   255
            Left            =   260
            TabIndex        =   27
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "  "
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
         Begin Threed.SSPanel SSPanel14 
            Height          =   255
            Left            =   870
            TabIndex        =   28
            Top             =   130
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "  "
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BevelOuter      =   0
         End
      End
   End
   Begin VB.PictureBox pbxTop 
      Align           =   1  'Align Top
      Height          =   2640
      Left            =   0
      ScaleHeight     =   2580
      ScaleWidth      =   15270
      TabIndex        =   32
      Top             =   600
      Width           =   15330
      Begin VB.TextBox txtDbgElapsed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   10320
         MaxLength       =   25
         TabIndex        =   163
         Text            =   "debug"
         Top             =   2280
         Visible         =   0   'False
         Width           =   4890
      End
      Begin VB.TextBox txtEndOp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12480
         MaxLength       =   25
         TabIndex        =   152
         ToolTipText     =   "Alphanumeric Name"
         Top             =   2160
         Width           =   2190
      End
      Begin VB.Frame frmStnDtlMsg 
         Height          =   1200
         Left            =   4800
         TabIndex        =   139
         Top             =   480
         Width           =   6000
         Begin VB.TextBox txtDebug1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   240
            MaxLength       =   25
            TabIndex        =   153
            Text            =   "desc"
            Top             =   840
            Width           =   5490
         End
         Begin VB.TextBox txtStnDtlMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   945
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   140
            Text            =   "frmStnDe.frx":778A3
            Top             =   150
            Width           =   5800
         End
      End
      Begin VB.TextBox txtShift 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   8985
         TabIndex        =   138
         Text            =   "shift"
         Top             =   2295
         Width           =   415
      End
      Begin VB.TextBox txtDspShift 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   540
         Left            =   8985
         TabIndex        =   137
         Text            =   "8"
         Top             =   1755
         Width           =   415
      End
      Begin VB.TextBox txtStation 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   5985
         TabIndex        =   136
         Text            =   "station"
         Top             =   2295
         Width           =   735
      End
      Begin VB.TextBox txtDspStn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   630
         Left            =   6150
         TabIndex        =   135
         Text            =   "8"
         Top             =   1710
         Width           =   415
      End
      Begin VB.CommandButton cmdDnShift 
         DisabledPicture =   "frmStnDe.frx":778AD
         DownPicture     =   "frmStnDe.frx":77FAF
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   8100
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmStnDe.frx":786B1
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdUpShift 
         DisabledPicture =   "frmStnDe.frx":78DB3
         DownPicture     =   "frmStnDe.frx":794B5
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   9465
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmStnDe.frx":79BB7
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdDnStn 
         DisabledPicture =   "frmStnDe.frx":7A2B9
         DownPicture     =   "frmStnDe.frx":7AEFB
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmStnDe.frx":7BB3D
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdUpStn 
         DisabledPicture =   "frmStnDe.frx":7C77F
         DownPicture     =   "frmStnDe.frx":7D3C1
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmStnDe.frx":7E003
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.TextBox txtEngineer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         MaxLength       =   25
         TabIndex        =   106
         ToolTipText     =   "Alphanumeric Name"
         Top             =   2130
         Width           =   2760
      End
      Begin VB.TextBox txtVehicle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         MaxLength       =   25
         TabIndex        =   105
         Text            =   "1234567890123456789012345"
         ToolTipText     =   "Alphanumeric Vehicle Identification Number"
         Top             =   1845
         Width           =   2760
      End
      Begin VB.Frame frmStnRecipe 
         Caption         =   "Recipe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1200
         Left            =   0
         TabIndex        =   100
         Top             =   480
         Width           =   4800
         Begin VB.TextBox txtRecipeName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   120
            MaxLength       =   70
            TabIndex        =   104
            Text            =   "desc"
            ToolTipText     =   "Station Recipe Name"
            Top             =   210
            Width           =   4575
         End
         Begin VB.TextBox txtRcpDsc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   360
            MaxLength       =   70
            TabIndex        =   103
            Text            =   "purge"
            Top             =   690
            Width           =   4320
         End
         Begin VB.TextBox txtRcpDsc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   360
            MaxLength       =   70
            TabIndex        =   102
            Text            =   "load"
            Top             =   915
            Width           =   4320
         End
         Begin VB.TextBox txtRcpDsc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            MaxLength       =   70
            TabIndex        =   101
            Text            =   "cycles"
            Top             =   450
            Width           =   4560
         End
      End
      Begin VB.Frame frmStnCanister 
         Caption         =   "Canister"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1200
         Left            =   10800
         TabIndex        =   92
         Top             =   480
         Width           =   4045
         Begin VB.TextBox txtLeakCheckStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   60
            MaxLength       =   25
            TabIndex        =   151
            Text            =   "leakcheck results"
            ToolTipText     =   "Station Canister Description"
            Top             =   450
            Width           =   3930
         End
         Begin VB.TextBox txtCanID 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   120
            MaxLength       =   25
            TabIndex        =   93
            Text            =   "desc"
            ToolTipText     =   "Station Canister Description"
            Top             =   210
            Width           =   3815
         End
         Begin VB.Label Label5 
            Caption         =   "Canister Work Cap:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   99
            Top             =   915
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Canister Volume:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   98
            Top             =   690
            Width           =   2055
         End
         Begin VB.Label lblCanVolUnits 
            Caption         =   "liters"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3240
            TabIndex        =   97
            Top             =   690
            Width           =   615
         End
         Begin VB.Label lblCanWcUnits 
            Caption         =   "grams"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3240
            TabIndex        =   96
            Top             =   915
            Width           =   675
         End
         Begin VB.Label lblBedVolume 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2280
            TabIndex        =   95
            ToolTipText     =   "Current Canister Size in liters"
            Top             =   690
            Width           =   855
         End
         Begin VB.Label lblWorkCap 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2280
            TabIndex        =   94
            ToolTipText     =   "Current Canister Size in grams"
            Top             =   915
            Width           =   855
         End
      End
      Begin VB.TextBox txtStartOp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12480
         MaxLength       =   25
         TabIndex        =   91
         ToolTipText     =   "Alphanumeric Name"
         Top             =   1845
         Width           =   2190
      End
      Begin Threed.SSPanel pnlStatusFrame 
         Height          =   480
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   4800
         _Version        =   65536
         _ExtentX        =   8467
         _ExtentY        =   847
         _StockProps     =   15
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodColor      =   16777088
         Begin Threed.SSPanel pnlStatus 
            Height          =   300
            Left            =   90
            TabIndex        =   57
            ToolTipText     =   "Station Status"
            Top             =   90
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Station Status"
            ForeColor       =   -2147483625
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel pnlReportFrame 
         Height          =   480
         Left            =   10800
         TabIndex        =   58
         Top             =   0
         Width           =   4045
         _Version        =   65536
         _ExtentX        =   7135
         _ExtentY        =   847
         _StockProps     =   15
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodColor      =   16777088
         Begin Threed.SSPanel pnlReport 
            Height          =   300
            Left            =   90
            TabIndex        =   59
            ToolTipText     =   "Station Report Number"
            Top             =   90
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Station Report Number"
            ForeColor       =   -2147483625
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel pnlNameFrame 
         Height          =   480
         Left            =   4800
         TabIndex        =   61
         Top             =   0
         Width           =   6000
         _Version        =   65536
         _ExtentX        =   10583
         _ExtentY        =   847
         _StockProps     =   15
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodColor      =   16777088
         Begin Threed.SSPanel pnlStnName 
            Height          =   300
            Left            =   90
            TabIndex        =   74
            ToolTipText     =   "Station Number"
            Top             =   90
            Width           =   5820
            _Version        =   65536
            _ExtentX        =   10266
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Station Name"
            ForeColor       =   -2147483625
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
      End
      Begin VB.Label lblVehicle 
         Caption         =   "Vehicle No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   110
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label lblEngineer 
         Caption         =   "Engineer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   109
         Top             =   2145
         Width           =   1335
      End
      Begin VB.Label lblEndOp 
         Caption         =   "End Operator:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   108
         Top             =   2145
         Width           =   1455
      End
      Begin VB.Label lblStartOp 
         Caption         =   "Start Operator:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   107
         Top             =   1860
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar tbrNavigate 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1058
      ButtonWidth     =   1058
      ButtonHeight    =   953
      _Version        =   393216
   End
   Begin Threed.SSPanel SSPanel9 
      Height          =   375
      Left            =   0
      TabIndex        =   141
      ToolTipText     =   "Purge Differential Pressure Reading"
      Top             =   0
      Width           =   2925
      _Version        =   65536
      _ExtentX        =   5159
      _ExtentY        =   661
      _StockProps     =   15
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   8
      Begin Threed.SSPanel SSPanel17 
         Height          =   345
         Left            =   -1920
         TabIndex        =   142
         Top             =   120
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "1"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel SSPanel18 
         Height          =   285
         Left            =   1120
         TabIndex        =   143
         ToolTipText     =   "Current Purge Differential Pressure Reading"
         Top             =   60
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge DP"
         Height          =   285
         Left            =   120
         TabIndex        =   145
         Top             =   90
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "in H2O"
         Height          =   285
         Left            =   2200
         TabIndex        =   144
         Top             =   90
         Width           =   1035
      End
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logou&t"
      End
      Begin VB.Menu mnuCopyFile 
         Caption         =   "&Copy Files"
      End
      Begin VB.Menu mnuPrintFile 
         Caption         =   "&Print Files"
      End
      Begin VB.Menu beforeExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Program"
      End
   End
   Begin VB.Menu mnuEditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCanisters 
         Caption         =   "Ca&nisters"
      End
      Begin VB.Menu mnuRecipes 
         Caption         =   "&Recipes"
      End
      Begin VB.Menu mnuCourses 
         Caption         =   "Co&urses"
      End
      Begin VB.Menu mnuPurgeProfiles 
         Caption         =   "&PurgeProfiles"
      End
      Begin VB.Menu mnuConfiguration 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuSysdef 
         Caption         =   "&System Definition"
      End
      Begin VB.Menu mnuTomCanLoad 
         Caption         =   "&TOM Can Load"
      End
   End
   Begin VB.Menu mnuViewMenu 
      Caption         =   "&View"
      Begin VB.Menu mnuAirLog 
         Caption         =   "&AirLog"
      End
      Begin VB.Menu mnuButane 
         Caption         =   "&Butane Available"
      End
      Begin VB.Menu mnuEventLog 
         Caption         =   "&Event Log"
      End
      Begin VB.Menu mnuFuelUseLog 
         Caption         =   "&Fuel Consumption Log"
      End
      Begin VB.Menu mnuJoblist 
         Caption         =   "&Joblist"
      End
      Begin VB.Menu mnuOotMonitor 
         Caption         =   "&OOT Monitor"
      End
   End
   Begin VB.Menu mnuDataMenu 
      Caption         =   "&Data"
      Begin VB.Menu mnuReviewData 
         Caption         =   "&Review Data"
      End
      Begin VB.Menu mnuWatchData 
         Caption         =   "&Watch Current Data"
      End
   End
   Begin VB.Menu mnuToolsMenu 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalibration 
         Caption         =   "&Calibration"
      End
      Begin VB.Menu mnuIomonitor 
         Caption         =   "&I/O Monitor"
      End
      Begin VB.Menu mnuScaleMonitor 
         Caption         =   "&Scale Monitor"
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu mnuOperatorManual 
         Caption         =   "&Operator Manual"
      End
      Begin VB.Menu mnuFirstAid 
         Caption         =   "&FirstAid File Save"
      End
      Begin VB.Menu beforeAbout 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About CPS release7"
      End
   End
End
Attribute VB_Name = "frmStnDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 89 '''''''''''''''Form STNDETAIL.frm ''''''''''''''''''''
Option Explicit

Private stnDtl_StnMode_Last(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Private stnDtl_StnCourse_Last(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Private stnDtl_DispStn_Last As Integer
Private stnDtl_DispShift_Last As Integer
Private sStr As String
Private clipLoEU As Single
Private clipLoPercent As Single
Private clipHiPercent As Single
Private flagShowIt As Boolean
Private tempEU As Single
Private temptime As Date
Private tempDay As Long
Private tempHr As Long
Private tempHr2 As Long
Private tempMin As Long
Private tempSec As Long
Private pauseSec As Long
Private tempPercent As Single
Private tempPercentSP As Single
Private tempStn As Integer
Private tempFunc As Integer
Private tempFuncSP As Integer
Private tempText As String
Private tempVal As Double
Private grams As Variant
Private dSeconds As Double
Private WhichGraph As Integer
Private ShowXYLegend As Boolean
Private GraphsTop As Integer
Private ElapsedBoxTop As Integer
Private EstJobDurBoxTop As Integer
Private DelayBoxTop As Integer
Private PurgeDpBoxTop As Integer
Private ScaleBoxTop As Integer
Private CycleBoxTop As Integer
Private CourseBoxTop As Integer
Private WaterBathBoxTop As Integer
Private OversizeBoxTop As Integer
Private OversizeBoxLeft As Integer
Private OversizeBoxWidth As Integer
Private BoxLblLeft As Integer
Private BoxMaxLeft As Integer
Private BoxActLeft As Integer
Private ScreenDebug As Boolean
Const DefaultCommentHeight = 600
Const DefaultCommentFrameHeight = 870
Const TopOfCmdButtons = 375
Const ShowBarGraphs = 1
Const ShowXYGraphs = 2
' Note: the number of points displayed on the graph is the number of elements
' allocated in the first dimension of the Graph array
Private Graph(NumPoints, 1 To 6) As Single
Private NumPointsSoFar(1 To MAX_STN, 1 To MAX_SHIFT) As Single
Private StnGraph(1 To MAX_STN, 1 To MAX_SHIFT, NumPoints, 1 To 6) As Single

Sub Update_Stn(ByVal Index As Integer, ByVal Index2 As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 89, 1
Dim dMsg As String

    
' txtDbgElapsed.text = LoadControl(Index, index2).PhaseDts & " - " & LoadPhaseDesc(LoadControl(Index, index2).Phase)
' txtDbgElapsed.text = LoadPhaseDesc(LoadControl(Index, index2).Phase) & " - " & Format(LoadControl(Index, index2).TotalWtChg, "##0.000") & " - " & Format(LoadControl(Index, index2).ElapsedHours, "##0.000") & " - " & Format(LoadControl(Index, index2).TotalWtChgRate, "###0.00")
' txtDebug1.text = IIf(LoadControl(Index, index2).WaterBathTempOK, "WaterBath OK", "WaterBath NOT ok")
txtDebug1.Visible = False
    
    ' Update Text
    If StationControl(Index, Index2).Course <> stnDtl_StnCourse_Last(Index, Index2) Then UpdateStnRcpDsc Index, Index2
    If StationControl(Index, Index2).Course <> stnDtl_StnCourse_Last(Index, Index2) Then Update_Text Index, Index2
    If StationControl(Index, Index2).Mode <> stnDtl_StnMode_Last(Index, Index2) Then Update_Text Index, Index2
    ' Update the Navigate Toolbar buttons
    UpdateNavigateBtns
    
    ' Status Bars
    UpdateStatusBars
    
    '**************************************************************
    ChgErrModule 89, 2
    txtDspStn.text = Index
    txtDspShift.text = Index2
    ' Clock (PurgeAir) Panel
    pnlPurgeAir.ForeColor = frmMainMenu.pnlPurgeAir.ForeColor
    pnlPurgeAir.Caption = frmMainMenu.pnlPurgeAir.Caption
    
    ChgErrModule 89, 3
    ' **************************************************************************************
    ' MODE DISPLAY
    If StationControl(Index, Index2).Mode <> stnDtl_StnMode_Last(Index, Index2) _
        Or DispStn <> stnDtl_DispStn_Last _
        Or DispShift <> stnDtl_DispShift_Last _
        Or StationControl(Index, Index2).Mode = VBLEAK _
        Or StationControl(Index, Index2).Mode = VBLOAD _
        Or StationControl(Index, Index2).Mode = VBPURGE _
        Or StationControl(Index, Index2).Mode = VBPOSTLEAK _
        Or StationControl(Index, Index2).Mode = VBPOSTLOAD _
        Or StationControl(Index, Index2).Mode = VBPOSTPURGE _
        Or StationControl(Index, Index2).Mode = VBSCALEWAIT _
        Or StationControl(Index, Index2).Mode = VBSTARTWAIT Then
        
        ' only update if mode has changed (or description has a variable in it)
        Select Case StationControl(Index, Index2).Mode
            Case VBLEAK
                ' Leak Check - add leak check phase description
                tempText = ModeDescShort(VBLEAK) & " - " & LeakPhaseDesc(LeakCheckControl.Phase) & " " & LeakMethodDesc(LeakCheckControl.Method)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case VBLOAD
                ' Loading or Waiting for Scales to Settle?
                If LoadControl(Index, Index2).Phase = LoadPause Then
                    ' Waiting for Scales to Settle
                    tempText = " Load Settling for "
                    tempText = tempText & Format(StationConfig(Index, Index2).LoadSettleTime, "##0.0#")
                    tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                    tempText = tempText
                    tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                Else
                    ' Loading - add load method description
                    Select Case StationRecipe(Index, Index2).Load_MethodSave
                        Case NOLOAD
                            tempText = LoadTypeDesc(NOLOAD)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc2(NOLOAD)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(NOLOAD)
                        Case LOADBYTIME
                            tempText = LoadTypeDesc(LOADBYTIME)
                            tempText = tempText & Format(StationRecipe(Index, Index2).Load_Time, "##0")
                            tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                        Case LOADBYWC
                            tempText = LoadTypeDesc(LOADBYWC)
                            tempText = tempText & Format(StationRecipe(Index, Index2).WC_MultSave, "##0.#")
                            tempText = tempText & LoadTypeDesc2(LOADBYWC)
                            tempText = tempText & Format(StationRecipe(Index, Index2).EPAFill, "##0")
                            tempText = tempText & LoadTypeDesc3(LOADBYWC)
                        Case LOADBYWEIGHT
                            tempText = LoadTypeDesc(LOADBYWEIGHT)
                            If Int(StationRecipe(Index, Index2).Load_Wt) = StationRecipe(Index, Index2).Load_Wt Then
                                ' no digits to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(Index, Index2).Load_Wt, "##0")
                            Else
                                ' digit(s) to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(Index, Index2).Load_Wt, "##0.##")
                            End If
                            tempText = tempText & LoadTypeDesc2(LOADBYWEIGHT)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYWEIGHT)
                        Case LOADBYBREAKTHRU
                            tempText = LoadTypeDesc(LOADBYBREAKTHRU)
                            If Int(StationRecipe(Index, Index2).LoadBreakthrough) = StationRecipe(Index, Index2).LoadBreakthrough Then
                                ' no digits to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(Index, Index2).LoadBreakthrough, "##0")
                            Else
                                ' digit(s) to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(Index, Index2).LoadBreakthrough, "##0.##")
                            End If
                            tempText = tempText & LoadTypeDesc2(LOADBYBREAKTHRU)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYBREAKTHRU)
                        Case LOADBYFID
                            tempText = LoadTypeDesc(LOADBYFID)
                            tempText = tempText & Format(StationRecipe(Index, Index2).FIDmg, "#####0")
                            tempText = tempText & LoadTypeDesc2(LOADBYFID)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYFID)
                    End Select
                End If
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case VBPURGE
                ' Purging or Waiting for Scales to Settle?
                If PurgeControl(Index, Index2).Phase = PurgePause Then
                    ' Waiting for Scales to Settle
                    tempText = " Purge Settling for "
                    tempText = tempText & Format(StationConfig(Index, Index2).PurgeSettleTime, "##0.0#")
                    tempText = tempText & LoadTypeDesc2(PURGEBYTIME)
                    tempText = tempText
                    tempText = tempText & LoadTypeDesc3(PURGEBYTIME)
                Else
                    ' Purge - add Purge method description
                    Select Case StationRecipe(Index, Index2).Purge_Method
                        Case NOPURGE
                            tempText = "No Purge"
                        Case PURGEBYTIME
                            tempText = ModeDescShort(VBPURGE) & " for " & StationRecipe(Index, Index2).Purge_Time & " Minute"
                            If StationRecipe(Index, Index2).Purge_Time > 1 Then tempText = tempText & "s"
                        Case PURGEBYVOLUME
                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(Index, Index2).Purge_Can_Vol & " Canister Volume"
                            If StationRecipe(Index, Index2).Purge_Can_Vol <> 1 Then tempText = tempText & "s"
                        Case PURGEAUXONLY
                            tempText = ModeDescShort(VBPURGE) & " Aux Can for " & StationRecipe(Index, Index2).Purge_AuxTime & " Minute"
                            If StationRecipe(Index, Index2).Purge_AuxTime > 1 Then tempText = tempText & "s"
                        Case PURGEBYPROFILE
                            tempText = ModeDescShort(VBPURGE) & " " & " by Profile"
                        Case PURGEBYWC
                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(Index, Index2).Purge_TargetWC & " % of Work Cap"
                        Case PURGETOTARGET
                            tempText = ModeDescShort(VBPURGE) & " to " & StationRecipe(Index, Index2).Purge_TargetWeight & " grams"
                        Case PURGETOUNDOLOAD
                            tempText = ModeDescShort(VBPURGE) & " to " & " Undo Load"
                        Case PURGEBYLITERS
                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(Index, Index2).Purge_Liters & " liter"
                            If StationRecipe(Index, Index2).Purge_Liters <> 1 Then tempText = tempText & "s"
                    End Select
                End If
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case VBPOSTLEAK
                ' Post LeakCheck Pause
                tempText = ModeDescShort(VBPOSTLEAK)
                tempText = tempText & " for "
                tempText = tempText & Format(StationRecipe(Index, Index2).PauseLeakTime, "##0.0#")
                tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                tempText = tempText
                tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case VBPOSTLOAD
                ' Post Load Pause
                tempText = ModeDescShort(VBPOSTLOAD)
                tempText = tempText & " for "
                tempText = tempText & Format(StationRecipe(Index, Index2).PauseLoadTime, "##0.0#")
                tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                tempText = tempText
                tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case VBPOSTPURGE
                ' Post Purge Pause
                tempText = ModeDescShort(VBPOSTPURGE)
                tempText = tempText & " for "
                tempText = tempText & Format(StationRecipe(Index, Index2).PausePurgeTime, "##0.0#")
                tempText = tempText & LoadTypeDesc2(PURGEBYTIME)
                tempText = tempText
                tempText = tempText & LoadTypeDesc3(PURGEBYTIME)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case VBSCALEWAIT
                ' Waiting for Scale(s) - add which scale(s)
                tempText = ModeDescShort(VBSCALEWAIT)
                If StationRecipe(Index, Index2).UsePriScale And StationRecipe(Index, Index2).UseAuxScale Then
                    ' Using Two Scales
                    tempText = tempText & "s "
                    ' Scales in use ?
                    If Scale_In_Use(StationRecipe(Index, Index2).PriScaleNo) And Scale_In_Use(StationRecipe(Index, Index2).AuxScaleNo) Then
                        ' Both Scales in use
                        tempText = tempText & Format(StationRecipe(Index, Index2).PriScaleNo, "#0") & " && " & Format(StationRecipe(Index, Index2).AuxScaleNo, "#0")
                    ElseIf Scale_In_Use(StationRecipe(Index, Index2).PriScaleNo) Then
                        ' Primary Scale in use
                        tempText = tempText & Format(StationRecipe(Index, Index2).PriScaleNo, "#0")
                    ElseIf Scale_In_Use(StationRecipe(Index, Index2).AuxScaleNo) Then
                        ' Aux Scale in use
                        tempText = tempText & Format(StationRecipe(Index, Index2).AuxScaleNo, "#0")
                    End If
                ElseIf StationRecipe(Index, Index2).UsePriScale Then
                    ' Using Only Primary Scale
                    tempText = tempText & " "
                    If Scale_In_Use(StationRecipe(Index, Index2).PriScaleNo) Then tempText = tempText & Format(StationRecipe(Index, Index2).PriScaleNo, "#0")
                ElseIf StationRecipe(Index, Index2).UseAuxScale Then
                    ' Using Only Aux Scale
                    tempText = tempText & " "
                    If Scale_In_Use(StationRecipe(Index, Index2).AuxScaleNo) Then tempText = tempText & Format(StationRecipe(Index, Index2).AuxScaleNo, "#0")
                End If
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case VBSTARTWAIT
                ' Delayed Start - add how long
                tempText = StartTypeDesc(StationRecipe(Index, Index2).StartMethod)
                ' Which Method of Delay ?
                Select Case StationRecipe(Index, Index2).StartMethod
                    Case STARTNOW
                        tempText = tempText & StartTypeDesc2(STARTNOW)
                    Case STARTDELAYED
                        tempText = tempText & Format(StationRecipe(Index, Index2).StartDelay, "##0")
                        tempText = tempText & StartTypeDesc2(STARTDELAYED)
                    Case STARTATDATE
                        tempText = tempText & StartTypeDesc2(STARTATDATE)
                        tempText = tempText & Format(StationRecipe(Index, Index2).StartDate, "D MMM, YYYY   h:mm")
                End Select
                pnlStatus.Caption = tempText
                pnlStatus.ToolTipText = ""
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
            Case Else
                pnlStatus.Caption = ModeDescShort(StationControl(Index, Index2).Mode)
                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
        End Select
    End If
    
    ' JobNumber Panel
    If StationControl(Index, Index2).DBFile = "" Then
        pnlReport.Caption = "No Active Job File"
    Else
        pnlReport.Caption = "Job Number  " & StationControl(Index, Index2).Job_Number
    End If
    If pnlReport.BackColor <> pnlStatus.BackColor Then pnlReport.BackColor = pnlStatus.BackColor
    If pnlReport.ForeColor <> pnlStatus.ForeColor Then pnlReport.ForeColor = pnlStatus.ForeColor
    
    ' Station Name Panel
    If StationControl(Index, Index2).Mode <> stnDtl_StnMode_Last(Index, Index2) _
        Or DispStn <> stnDtl_DispStn_Last _
        Or DispShift <> stnDtl_DispShift_Last Then
            pnlStnName.BackColor = IIf(StationControl(Index, Index2).Mode = VBIDLE, pnlNameFrame.BackColor, pnlStatus.BackColor)
            pnlStnName.ForeColor = IIf(StationControl(Index, Index2).Mode = VBIDLE, pnlNameFrame.ForeColor, pnlStatus.ForeColor)
    End If
    ' **************************************************************************************
    
    
    
    ChgErrModule 89, 445
    ' **************************************************************************************
    ' Is This Shift the Active Shift for this Station
    If Stn_ActiveShift(Index) = Index2 Then
    
        GraphsTop = IIf((pbxRptName.Top = OutOfSight), (pbxBottom.Top - pnlBarGraphs.Height), OutOfSight)
        Select Case WhichGraph
            
            Case ShowXYGraphs
            
                ' Do Not Show Bar Graphs
                pnlBarGraphs.Top = OutOfSight
                ' Show XY Graphs
    '            pnlXYGraphs.Width = 5925
    '            chtStnChart = 5655
'                pnlXYGraphs.Width = 6905
'                chtStnChart.Width = 6725
'                pnlXYGraphs.Width = 14205
'                chtStnChart.Width = 14125
                pnlXYGraphs.Width = 11905
                chtStnChart.Width = 11585
                pnlXYGraphs.Top = GraphsTop
                        
            
            Case ShowBarGraphs
            
                ' Do Not Show XY Graphs
                pnlXYGraphs.Top = OutOfSight
                ' Show Bar Graphs
                pnlBarGraphs.Top = GraphsTop
                ' Update Bar Graph Descriptors
                If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And (StationRecipe(Index, Index2).LiveFuel)) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(Index, Index2).LiveFuel))) Then
                    pnlBarGraphs.Width = 5865 - pnlButane.Width
                    pnlNitrogen.Caption = "Vapor Carrier"
                Else
                    pnlBarGraphs.Width = 5865
                    pnlNitrogen.Caption = "Nitrogen"
                End If
                
                
                ' ********************************************************
                ' Update Bar Graph Displays
                clipHiPercent = 99.9
                clipLoPercent = 0.9
                
                       
                ' NITROGEN
                ChgErrModule 89, 446
                If (STN_INFO(Index).Type <> STN_DUMMY_TYPE) Then
                    
                    pnlNitCV.FloodColor = ModeBackColor(StationControl(Index, Index2).Mode)
                    pnlNitPV.FloodColor = BarActual_ForeColor
                    Select Case STN_INFO(Index).Type
                        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                            tempFunc = asNitrogenFlow
                            tempFuncSP = asNitrogenFlowSP
                        Case STN_ORVR2_TYPE
                            If StationRecipe(Index, Index2).UseHiRangeMFC Then
                                ' use higher range MFC
                                tempFunc = asNitrogenORVRFlow
                                tempFuncSP = asNitrogenORVRFlowSP
                            Else
                                ' use lower range MFC
                                tempFunc = asNitrogenFlow
                                tempFuncSP = asNitrogenFlowSP
                            End If
                        Case STN_LIVEFUEL_TYPE
                            tempFunc = asLiveFuelVaporFlow
                            tempFuncSP = asLiveFuelVaporFlowSP
                        Case STN_LIVEREG_TYPE
                            If StationRecipe(Index, Index2).LiveFuel Then
                                ' use Live Fuel
                                tempFunc = asLiveFuelVaporFlow
                                tempFuncSP = asLiveFuelVaporFlowSP
                            Else
                                ' use Butane/Nitrogen
                                tempFunc = asNitrogenFlow
                                tempFuncSP = asNitrogenFlowSP
                            End If
                        Case STN_LIVEORVR2_TYPE
                            If StationRecipe(Index, Index2).LiveFuel Then
                                ' use Live Fuel
                                If StationRecipe(Index, Index2).UseHiRangeMFC Then
                                    ' use higher range MFC
                                    tempFunc = asLiveFuelVaporORVRFlow
                                    tempFuncSP = asLiveFuelVaporORVRFlowSP
                                Else
                                    ' use lower range MFC
                                    tempFunc = asLiveFuelVaporFlow
                                    tempFuncSP = asLiveFuelVaporFlowSP
                                End If
                            Else
                                ' use Butane/Nitrogen
                                If StationRecipe(Index, Index2).UseHiRangeMFC Then
                                    ' use higher range MFC
                                    tempFunc = asNitrogenORVRFlow
                                    tempFuncSP = asNitrogenORVRFlowSP
                                Else
                                    ' use lower range MFC
                                    tempFunc = asNitrogenFlow
                                    tempFuncSP = asNitrogenFlowSP
                                End If
                            End If
                        Case STN_COMBO3_TYPE
                            ' future
                        Case Else
                            ' Do Nothing
                    End Select
                    
                    tempPercentSP = 0
                    tempPercent = 0
                    
                    If Stn_AIO(Index, tempFuncSP).EuMax <> 0 Then
                        tempPercentSP = (Stn_AIO(Index, tempFuncSP).EUValue / Stn_AIO(Index, tempFuncSP).EuMax) * 100
                        tempPercentSP = IIf(tempPercentSP > clipHiPercent, 100, tempPercentSP)
                        tempPercentSP = IIf(tempPercentSP < clipLoPercent, 0, tempPercentSP)
                    End If
                        
                    If Stn_AIO(Index, tempFunc).EuMax <> 0 Then
                        tempPercent = (Stn_AIO(Index, tempFunc).EUValue / Stn_AIO(Index, tempFunc).EuMax) * 100
                        tempPercent = IIf(tempPercent > clipHiPercent, 100, tempPercent)
                        tempPercent = IIf(tempPercent < clipLoPercent, 0, tempPercent)
                    End If
                    
                    pnlNitCV.FloodPercent = tempPercentSP
                    pnlNitPV.FloodPercent = tempPercent
                
                Else
                    ' No Nitrogen on Dummy Stations
                    pnlNitCV.FloodPercent = 0
                    pnlNitPV.FloodPercent = 0
                End If
                    
                
                ' BUTANE
                ChgErrModule 89, 447
                If (STN_INFO(Index).Type <> STN_LIVEFUEL_TYPE And STN_INFO(Index).Type <> STN_DUMMY_TYPE And (Not StationRecipe(Index, Index2).LiveFuel)) Then
                    
                    pnlBtnCV.FloodColor = ModeBackColor(StationControl(Index, Index2).Mode)
                    pnlBtnPV.FloodColor = BarActual_ForeColor
                    Select Case STN_INFO(Index).Type
                        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                            tempFunc = asButaneFlow
                            tempFuncSP = asButaneFlowSP
                        Case STN_ORVR2_TYPE
                            If StationRecipe(Index, Index2).UseHiRangeMFC Then
                                ' use higher range MFC
                                tempFunc = asButaneORVRFlow
                                tempFuncSP = asButaneORVRFlowSP
                            Else
                                ' use lower range MFC
                                tempFunc = asButaneFlow
                                tempFuncSP = asButaneFlowSP
                            End If
                        Case STN_LIVEFUEL_TYPE
                            ' Do Nothing
                        Case STN_LIVEREG_TYPE
                            ' using Butane/Nitrogen
                            tempFunc = asButaneFlow
                            tempFuncSP = asButaneFlowSP
                        Case STN_LIVEORVR2_TYPE
                            If StationRecipe(Index, Index2).UseHiRangeMFC Then
                                ' use higher range MFC
                                tempFunc = asButaneORVRFlow
                                tempFuncSP = asButaneORVRFlowSP
                            Else
                                ' use lower range MFC
                                tempFunc = asButaneFlow
                                tempFuncSP = asButaneFlowSP
                            End If
                        Case STN_COMBO3_TYPE
                            ' future
                        Case Else
                            ' Do Nothing
                    End Select
                    
                    tempPercentSP = 0
                    tempPercent = 0
                    
                    If Stn_AIO(Index, tempFuncSP).EuMax <> 0 Then
                        tempPercentSP = (Stn_AIO(Index, tempFuncSP).EUValue / Stn_AIO(Index, tempFuncSP).EuMax) * 100
                        tempPercentSP = IIf(tempPercentSP > clipHiPercent, 100, tempPercentSP)
                        tempPercentSP = IIf(tempPercentSP < clipLoPercent, 0, tempPercentSP)
                    End If
                        
                    If Stn_AIO(Index, tempFunc).EuMax <> 0 Then
                        tempPercent = (Stn_AIO(Index, tempFunc).EUValue / Stn_AIO(Index, tempFunc).EuMax) * 100
                        tempPercent = IIf(tempPercent > clipHiPercent, 100, tempPercent)
                        tempPercent = IIf(tempPercent < clipLoPercent, 0, tempPercent)
                    End If
                    
                    pnlBtnCV.FloodPercent = tempPercentSP
                    pnlBtnPV.FloodPercent = tempPercent
                
                Else
                    ' No Butane on Dummy or LiveFuel Stations
                    pnlBtnCV.FloodPercent = 0
                    pnlBtnPV.FloodPercent = 0
                End If
                    
                    
                ' PURGE AIR
                ChgErrModule 89, 442
                If (STN_INFO(Index).Type <> STN_DUMMY_TYPE) Then
                    
                    pnlPurCV.FloodColor = ModeBackColor(StationControl(Index, Index2).Mode)
                    pnlPurPV.FloodColor = BarActual_ForeColor
                    Select Case STN_INFO(Index).Type
                        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                            tempFunc = asPurgeAirFlow
                            tempFuncSP = asPurgeAirFlowSP
                        Case STN_ORVR2_TYPE
                            tempFunc = asPurgeAirFlow
                            tempFuncSP = asPurgeAirFlowSP
                        Case STN_LIVEFUEL_TYPE
                            tempFunc = asPurgeAirFlow
                            tempFuncSP = asPurgeAirFlowSP
                        Case STN_LIVEREG_TYPE
                            tempFunc = asPurgeAirFlow
                            tempFuncSP = asPurgeAirFlowSP
                        Case STN_LIVEORVR2_TYPE
                            tempFunc = asPurgeAirFlow
                            tempFuncSP = asPurgeAirFlowSP
                        Case STN_COMBO3_TYPE
                            ' future
                        Case Else
                            ' Do Nothing
                    End Select
                    
                    tempPercentSP = 0
                    tempPercent = 0
                    
                    If Stn_AIO(Index, tempFuncSP).EuMax <> 0 Then
                        tempPercentSP = (Stn_AIO(Index, tempFuncSP).EUValue / Stn_AIO(Index, tempFuncSP).EuMax) * 100
                        tempPercentSP = IIf(tempPercentSP > clipHiPercent, 100, tempPercentSP)
                        tempPercentSP = IIf(tempPercentSP < clipLoPercent, 0, tempPercentSP)
                    End If
                        
                    If Stn_AIO(Index, tempFunc).EuMax <> 0 Then
                        tempPercent = (Stn_AIO(Index, tempFunc).EUValue / Stn_AIO(Index, tempFunc).EuMax) * 100
                        tempPercent = IIf(tempPercent > clipHiPercent, 100, tempPercent)
                        tempPercent = IIf(tempPercent < clipLoPercent, 0, tempPercent)
                    End If
                    
                    pnlPurCV.FloodPercent = tempPercentSP
                    pnlPurPV.FloodPercent = tempPercent
                
                Else
                    ' No Purge Air on Dummy Stations
                    pnlPurCV.FloodPercent = 0
                    pnlPurPV.FloodPercent = 0
                End If
                
                    
                ' TARGET vs ACTUAL
                ChgErrModule 89, 5
                pnlTarget.FloodColor = ModeBackColor(StationControl(Index, Index2).Mode)
                pnlTarget.FloodPercent = IIf(StationControl(Index, Index2).Target > 0, 80, 0)          ' at 80 percent
                pnlActual.FloodColor = BarActual_ForeColor
                
                If StationControl(Index, Index2).Actual = 0 Or StationControl(Index, Index2).Target = 0 Then
                    tempPercent = 0
                Else
                    tempPercent = 80 * (StationControl(Index, Index2).Actual) / (StationControl(Index, Index2).Target)
                End If
                
                ChgErrModule 89, 11
                If StationControl(Index, Index2).Actual < 0 Then
                   tempPercent = 0
                Else
                    ' purge mode set values
                    If StationControl(Index, Index2).Mode = VBPURGE Then
                        If StationControl(Index, Index2).Actual > 0 Then
                            If StationControl(Index, Index2).Target > VALUE0 Then
                                tempPercent = ((StationControl(Index, Index2).Actual) / (StationControl(Index, Index2).Target)) * 80
                            End If
                        Else
                            pnlActual.FloodPercent = 0
                        End If
                    End If
                End If
                
                 
                ' Leak check set values
                ChgErrModule 89, 111
                If StationControl(Index, Index2).Mode = VBLEAK Then  ' leak check
                    tempPercent = 80 * (StationControl(Index, Index2).Actual + 0.0000001) / (0.0000001 + StationControl(Index, Index2).Target + 0.000001)
                End If '
                
                ' Clip Actual to 0-100 %
                If StationControl(Index, Index2).Mode = VBIDLE Then tempPercent = 0
                tempPercent = IIf(tempPercent > 100, 100, tempPercent)
                tempPercent = IIf(tempPercent < 0, 0, tempPercent)
                pnlActual.FloodPercent = tempPercent
                
                    
                ' Update Digital Displays
                ChgErrModule 89, 1111
                clipLoEU = 0.001
                txtPurCV.ForeColor = pnlPurCV.FloodColor
                txtNitCV.ForeColor = pnlNitCV.FloodColor
                txtBtnCV.ForeColor = pnlBtnCV.FloodColor
                txtTarget.ForeColor = pnlTarget.FloodColor
                txtPurPV.ForeColor = BarActual_ForeColor
                txtNitPV.ForeColor = BarActual_ForeColor
                txtBtnPV.ForeColor = BarActual_ForeColor
                txtActual.ForeColor = BarActual_ForeColor
                Select Case STN_INFO(Index).Type
                    Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                        tempEU = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtNitCV.text = Format(tempEU, "##0.000")
                        
                        tempEU = Stn_AIO(Index, asNitrogenFlow).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtNitPV.text = Format(tempEU, "##0.000")
                    
                        tempEU = Stn_AIO(Index, asButaneFlowSP).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtBtnCV.text = Format(tempEU, "##0.000")
                    
                        tempEU = Stn_AIO(Index, asButaneFlow).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtBtnPV.text = Format(tempEU, "##0.000")
                    
                        tempEU = Stn_AIO(Index, asPurgeAirFlowSP).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtPurCV.text = Format(tempEU, "##0.000")
                    
                        tempEU = Stn_AIO(Index, asPurgeAirFlow).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtPurPV.text = Format(tempEU, "##0.000")
                    
                
                    Case STN_ORVR2_TYPE
                        If StationRecipe(Index, Index2).UseHiRangeMFC Then
                            ' use higher range MFCs
                            tempEU = Stn_AIO(Index, asNitrogenORVRFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitCV.text = Format(tempEU, "##0.000")
                            
                            tempEU = Stn_AIO(Index, asNitrogenORVRFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitPV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asButaneORVRFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtBtnCV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asButaneORVRFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtBtnPV.text = Format(tempEU, "##0.000")
                        Else
                            ' use lower range MFCs
                            tempEU = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitCV.text = Format(tempEU, "##0.000")
                            
                            tempEU = Stn_AIO(Index, asNitrogenFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitPV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asButaneFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtBtnCV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asButaneFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtBtnPV.text = Format(tempEU, "##0.000")
                        End If
                        
                        tempEU = Stn_AIO(Index, asPurgeAirFlowSP).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtPurCV.text = Format(tempEU, "##0.000")
                    
                        tempEU = Stn_AIO(Index, asPurgeAirFlow).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtPurPV.text = Format(tempEU, "##0.000")
                       
                    
                    Case STN_LIVEFUEL_TYPE
                        tempEU = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtNitCV.text = Format(tempEU, "##0.000")
                        
                        tempEU = Stn_AIO(Index, asLiveFuelVaporFlow).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtNitPV.text = Format(tempEU, "##0.000")
                    
                        tempEU = Stn_AIO(Index, asPurgeAirFlowSP).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtPurCV.text = Format(tempEU, "##0.000")
                    
                        tempEU = Stn_AIO(Index, asPurgeAirFlow).EUValue
                        tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                        txtPurPV.text = Format(tempEU, "##0.000")
                        
                    Case STN_LIVEREG_TYPE
                        If StationRecipe(Index, Index2).LiveFuel Then
                            ' LIVE FUEL
                            tempEU = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitCV.text = Format(tempEU, "##0.000")
                            
                            tempEU = Stn_AIO(Index, asLiveFuelVaporFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitPV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asPurgeAirFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurCV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asPurgeAirFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurPV.text = Format(tempEU, "##0.000")
                            
                        Else
                            ' "REGULAR"
                            tempEU = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitCV.text = Format(tempEU, "##0.000")
                            
                            tempEU = Stn_AIO(Index, asNitrogenFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtNitPV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asButaneFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtBtnCV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asButaneFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtBtnPV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asPurgeAirFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurCV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asPurgeAirFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurPV.text = Format(tempEU, "##0.000")
                            
                        End If
                    
                    Case STN_LIVEORVR2_TYPE
                        If StationRecipe(Index, Index2).LiveFuel Then
                            ' LIVE FUEL
                            If StationRecipe(Index, Index2).UseHiRangeMFC Then
                                ' use higher range MFC
                                tempEU = Stn_AIO(Index, asLiveFuelVaporORVRFlowSP).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitCV.text = Format(tempEU, "##0.000")
                                
                                tempEU = Stn_AIO(Index, asLiveFuelVaporORVRFlow).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitPV.text = Format(tempEU, "##0.000")
                            Else
                                ' use lower range MFC
                                tempEU = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitCV.text = Format(tempEU, "##0.000")
                                
                                tempEU = Stn_AIO(Index, asLiveFuelVaporFlow).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitPV.text = Format(tempEU, "##0.000")
                            End If
                            
                            tempEU = Stn_AIO(Index, asPurgeAirFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurCV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asPurgeAirFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurPV.text = Format(tempEU, "##0.000")
                            
                        Else
                            ' REGULAR
                            If StationRecipe(Index, Index2).UseHiRangeMFC Then
                                ' use higher range MFCs
                                tempEU = Stn_AIO(Index, asNitrogenORVRFlowSP).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitCV.text = Format(tempEU, "##0.000")
                                
                                tempEU = Stn_AIO(Index, asNitrogenORVRFlow).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitPV.text = Format(tempEU, "##0.000")
                            
                                tempEU = Stn_AIO(Index, asButaneORVRFlowSP).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtBtnCV.text = Format(tempEU, "##0.000")
                            
                                tempEU = Stn_AIO(Index, asButaneORVRFlow).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtBtnPV.text = Format(tempEU, "##0.000")
                            Else
                                ' use lower range MFCs
                                tempEU = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitCV.text = Format(tempEU, "##0.000")
                                
                                tempEU = Stn_AIO(Index, asNitrogenFlow).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtNitPV.text = Format(tempEU, "##0.000")
                            
                                tempEU = Stn_AIO(Index, asButaneFlowSP).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtBtnCV.text = Format(tempEU, "##0.000")
                            
                                tempEU = Stn_AIO(Index, asButaneFlow).EUValue
                                tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                                txtBtnPV.text = Format(tempEU, "##0.000")
                            End If
                            
                            tempEU = Stn_AIO(Index, asPurgeAirFlowSP).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurCV.text = Format(tempEU, "##0.000")
                        
                            tempEU = Stn_AIO(Index, asPurgeAirFlow).EUValue
                            tempEU = IIf(tempEU < clipLoEU, 0, tempEU)
                            txtPurPV.text = Format(tempEU, "##0.000")
                                
                        End If
                    
                    Case STN_COMBO3_TYPE
                        ' future
                        
                     Case Else
                    ' Do Nothing
                    
                End Select
                
                txtTarget.text = Format(StationControl(Index, Index2).Target, "##0.00")
                txtActual.text = Format(StationControl(Index, Index2).Actual, "##0.00")
                
                ' Set Display for Out of Tolerance
                pnlNitrogen.ForeColor = IIf(OOTs(Index, Index2).NitFlowOOT, MEDRED, Black)
                pnlButane.ForeColor = IIf(OOTs(Index, Index2).BtnFlowOOT, MEDRED, Black)
                pnlPurge.ForeColor = IIf(OOTs(Index, Index2).PurFlowOOT, MEDRED, Black)
            
        End Select
    Else
    
        ' DO Not Show Bar Graphs
        pnlBarGraphs.Top = OutOfSight
        If WhichGraph = ShowXYGraphs Then
                ' Show XY Graphs
    '            pnlXYGraphs.Width = 5925
    '            chtStnChart = 5655
'                pnlXYGraphs.Width = 6905
'                chtStnChart.Width = 6725
                pnlXYGraphs.Width = 11905
                chtStnChart.Width = 11585
                pnlXYGraphs.Top = GraphsTop
        Else
            ' DO Not Show XY Graphs
            pnlXYGraphs.Top = OutOfSight
        End If
        
    End If
    ' **************************************************************************************
    
    
    
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ' Vertical Status Panel
    ' Vertical Status Panel
    ' Vertical Status Panel
    ' Vertical Status Panel
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
'    pbxVertical.Top = IIf((pbxRptName.Top = OutOfSight), GraphsTop, OutOfSight)
'    pbxVertical.Height = pnlBarGraphs.Height
    
    '**********************
    '**********************
    '**********************
    ' Elapsed TestTime Display & Job Duration Display
    ' Elapsed TestTime Display & Job Duration Display
    ' Elapsed TestTime Display & Job Duration Display
    '   NOTE:  86,400sec/day; 31,535,975sec/year
    '**********************
    '**********************
    '**********************
    If StationControl(Index, Index2).Mode = VBIDLE Then
        pnlTestTime.Top = OutOfSight
        pnlJobDuration.Top = OutOfSight
    Else
        ' Elapsed Test Time
        pnlTestTime.Top = ElapsedBoxTop
        pnlTestTime.Caption = "Elapsed:  " & DurationDescription(CSng(StationControl(Index, Index2).TestTimer) / CSng(60))
        pnlTestTime.ToolTipText = "Elapsed Time since Job Began"
        
        If StationControl(Index, Index2).TestTimerIsRunning Then
            If pnlTestTime.ForeColor <> pnlCycle.ForeColor Then pnlTestTime.ForeColor = pnlCycle.ForeColor
        Else
            If pnlTestTime.ForeColor <> Black Then pnlTestTime.ForeColor = Black
        End If
        
        ' Estimated Job Duration
        pnlJobDuration.Top = EstJobDurBoxTop
        pnlJobDuration.Caption = "Estimate: " & StationControl(Index, Index2).EstJobDurDesc
        pnlJobDuration.ForeColor = pnlTestTime.ForeColor
    End If
    
    '**********************
    '**********************
    '**********************
    ' Scale Weight Display
    ' Scale Weight Display
    ' Scale Weight Display
    '**********************
    '**********************
    '**********************
    If StationControl(Index, Index2).Mode = VBIDLE Then
        pnlWtPri.Caption = " "
        pnlWtAux.Caption = " "
        pnlChgPri.Caption = " "
        pnlChgAux.Caption = " "
        lblAuxScale.Caption = "Aux."
        lblPriScale.Caption = "Pri."
        pnlScale.Top = OutOfSight
    Else
        If StationRecipe(Index, Index2).UsePriScale = False And StationRecipe(Index, Index2).UseAuxScale = False Then
            pnlWtPri.Caption = " "
            pnlWtAux.Caption = " "
            pnlChgPri.Caption = " "
            pnlChgAux.Caption = " "
            lblAuxScale.Caption = "Aux."
            lblPriScale.Caption = "Pri."
            pnlScale.Top = OutOfSight
        Else
            pnlScale.Top = ScaleBoxTop
            If StationRecipe(Index, Index2).UsePriScale = True Then
                pnlWtPri.Caption = Format(StationControl(Index, Index2).PriScaleWt, "###0.00")
'                pnlChgPri.ForeColor = ModeBackColor(StationControl(Index, index2).Mode)
                Select Case StationControl(Index, Index2).Mode
                    Case VBLOAD, VBPOSTLOAD
                        pnlChgPri.Caption = Format(LoadControl(Index, Index2).PriWtChg, "###0.00")
                    Case VBPURGE, VBPOSTPURGE
                        pnlChgPri.Caption = Format(PurgeControl(Index, Index2).PriWtChg, "###0.00")
                    Case VBPURGEWAIT, VBPRELOAD, VBCOMPLETE
                        pnlChgPri.Caption = " "
                    Case VBPAUSE, VBPAUSEBYUSER, VBPAUSEOOT, VBPAUSEALARM, VBGASPAUSE, VBWBPAUSE
                        Select Case StationControl(Index, Index2).Mode_PauseSave
                            Case VBLOAD, VBPOSTLOAD
                                pnlChgPri.Caption = Format(LoadControl(Index, Index2).PriWtChg, "###0.00")
                            Case VBPURGE, VBPOSTPURGE
                                pnlChgPri.Caption = Format(PurgeControl(Index, Index2).PriWtChg, "###0.00")
                            Case VBPURGEWAIT, VBPRELOAD, VBCOMPLETE
                                pnlChgPri.Caption = " "
                            Case Else
                                ' no change
                        End Select
                    Case Else
                        ' no change
                End Select
                pnlWtPri.Top = 90
                pnlChgPri.Top = pnlWtPri.Top
                lblPriScale.Top = pnlWtPri.Top + 75
                lblPriScale.Caption = "Pri. #" & Format(StationRecipe(Index, Index2).PriScaleNo, "#0")
                lblPriScale.Width = 735
            Else
                lblPriScale.Caption = "No Primary"
                lblPriScale.Width = 1600
                pnlWtPri.Caption = ""
                pnlWtPri.Top = OutOfSight
                pnlChgPri.Caption = ""
                pnlChgPri.Top = OutOfSight
            End If
            If StationRecipe(Index, Index2).UseAuxScale = True Then
                pnlWtAux.Caption = Format(StationControl(Index, Index2).AuxScaleWt, "###0.00")
'                pnlChgAux.ForeColor = ModeBackColor(StationControl(Index, index2).Mode)
                Select Case StationControl(Index, Index2).Mode
                    Case VBLOAD
                        pnlChgAux.Caption = Format(LoadControl(Index, Index2).AuxWtChg, "###0.00")
                    Case VBPURGE
                        pnlChgAux.Caption = Format(PurgeControl(Index, Index2).AuxWtChg, "###0.00")
                    Case VBPURGEWAIT, VBPRELOAD, VBCOMPLETE
                        pnlChgAux.Caption = " "
                    Case Else
                        ' no change
                End Select
                pnlWtAux.Top = 465
                pnlChgAux.Top = pnlWtAux.Top
                lblAuxScale.Top = pnlWtAux.Top + 75
                lblAuxScale.Caption = "Aux. #" & Format(StationRecipe(Index, Index2).AuxScaleNo, "#0")
                lblAuxScale.Width = 735
            Else
                lblAuxScale.Caption = "No Aux Scale"
                lblAuxScale.Width = 1600
                pnlWtAux.Caption = ""
                pnlWtAux.Top = OutOfSight
                pnlChgAux.Caption = ""
                pnlChgAux.Top = OutOfSight
            End If
        End If
    End If
        
    '**********************
    '**********************
    '**********************
    '  Purge Profile Panel
    '  Purge Profile Panel
    '  Purge Profile Panel
    '**********************
    '**********************
    '**********************
    pnlStepMax.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
    pnlMinutesMax.ForeColor = pnlStepMax.ForeColor
    Select Case StationControl(Index, Index2).Mode
           
        Case VBPURGE
            If (PurgeControl(Index, Index2).Phase = PurgePurging) Then
    
                Select Case StationRecipe(Index, Index2).Purge_Method
                    Case PURGEBYPROFILE
                        ' Purging by Profile
                        pnlProfile.Top = DelayBoxTop
                        pnlProfile.Caption = "Purging By Profile"
                        lblProfileStep.Left = BoxLblLeft
                        pnlStepActual.Left = BoxActLeft
                        pnlStepMax.Left = BoxMaxLeft
                        lblStepMinutes.Left = BoxLblLeft
                        pnlMinutesActual.Left = BoxActLeft
                        pnlMinutesMax.Left = BoxMaxLeft
                        pnlProfile.Caption = "Purge Profile"
                        lblProfileStep.Caption = "max/step"
                        lblStepMinutes.Caption = "max/minutes"
                        pnlStepActual.Caption = Format(PurgeControl(Index, Index2).curStep, "##0")
                        pnlStepMax.Caption = Format(StationProfile(Index, Index2).EndStep, "##0")
                        pnlMinutesActual.Caption = Format(PurgeControl(Index, Index2).StepElapsedMinutes, "###0.00")
                        pnlMinutesMax.Caption = Format(StationProfile(Index, Index2).StepDuration(PurgeControl(Index, Index2).curStep), "###0.00")
                    Case PURGEBYLITERS
                        ' Purging By Liters
                        pnlProfile.Top = DelayBoxTop
                        lblProfileStep.Left = BoxLblLeft
                        pnlStepActual.Left = BoxActLeft
                        pnlStepMax.Left = BoxMaxLeft
                        pnlMinutesMax.Left = OutOfSight
                        Select Case StationRecipe(Index, Index2).Purge_TargetMode
                            Case TARGETCONTINUOUS
                                lblProfileStep.Caption = "liters"
                                lblStepMinutes.Left = OutOfSight
                                pnlStepActual.Caption = Format(StationControl(Index, Index2).Actual, "###0.0")
                                pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0")
                                pnlMinutesActual.Left = OutOfSight
                                pnlProfile.Caption = "Purging By Liters"
                            Case TARGETPURGEPAUSE
                                lblProfileStep.Caption = "liters"
                                lblStepMinutes.Left = OutOfSight    'BoxLblLeft
                                lblStepMinutes.Caption = "step liters"
                                pnlStepActual.Caption = Format(StationControl(Index, Index2).Actual, "###0.0")
                                pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0.0")
                                pnlMinutesActual.Left = OutOfSight    'BoxActLeft
                                pnlMinutesActual.Caption = Format(PurgeControl(Index, Index2).Purge_Total, "#,##0.0")
                                Select Case PurgeControl(Index, Index2).curStep
                                    Case 0
                                        ' pausing
                                        sStr = "PurgePause #"
                                        sStr = sStr & Format(PurgeControl(Index, Index2).curCycle, "##0")
                                        sStr = sStr & " - " & Format((StationRecipe(Index, Index2).Purge_TargetPause - PurgeControl(Index, Index2).StepElapsedMinutes), "###0.00")
                                        sStr = sStr & " min to go"
                                    Case 1
                                        ' purging
                                        sStr = "Purging #"
                                        sStr = sStr & Format(PurgeControl(Index, Index2).curCycle, "##0")
                                        sStr = sStr & " - " & Format((StationRecipe(Index, Index2).Purge_TargetPurge - PurgeControl(Index, Index2).StepElapsedMinutes), "###0.00")
                                        sStr = sStr & " min to go"
                                End Select
                                pnlProfile.Caption = sStr
                        End Select
                        
                    Case PURGEBYVOLUME, PURGEBYWC, PURGETOTARGET, PURGETOUNDOLOAD
                        ' Purging By Volume & Purging By WC & Purging to Target & Purging to UndoLoad
                        pnlProfile.Top = DelayBoxTop
                        lblProfileStep.Left = BoxLblLeft
                        pnlStepActual.Left = BoxActLeft
                        pnlStepMax.Left = BoxMaxLeft
                        ' Purge By Volume doesn't display second row (Vol & target Vol)
                        If (StationRecipe(Index, Index2).Purge_Method = PURGEBYVOLUME) Then
                            lblStepMinutes.Left = OutOfSight
                            pnlMinutesActual.Left = OutOfSight
                            pnlMinutesMax.Left = OutOfSight
                        Else
                            lblStepMinutes.Left = BoxLblLeft
                            pnlMinutesActual.Left = BoxActLeft
                            pnlMinutesMax.Left = BoxMaxLeft
                        End If
                        Select Case StationRecipe(Index, Index2).Purge_TargetMode
                            Case TARGETCONTINUOUS
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYVOLUME) Then lblProfileStep.Caption = "Volumes"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYWC) Then lblProfileStep.Caption = "% Work Cap"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOTARGET) Then lblProfileStep.Caption = "Grams"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOUNDOLOAD) Then lblProfileStep.Caption = "Grams"
                                lblStepMinutes.Caption = "Volumes"
                                pnlStepActual.Caption = Format(StationControl(Index, Index2).Actual, "###0.0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYVOLUME) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYWC) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOTARGET) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0.0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOUNDOLOAD) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0.0")
                                pnlMinutesActual.Caption = Format(PurgeControl(Index, Index2).Purge_Volumes, "#,##0.0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYVOLUME) Then pnlProfile.Caption = "Purging By Volumes"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYWC) Then pnlProfile.Caption = "Purging By Working Capacity"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOTARGET) Then pnlProfile.Caption = "Purging To Target Weight"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOUNDOLOAD) Then pnlStepMax.Caption = pnlProfile.Caption = "Purging To Undo Load"
                            Case TARGETPURGEPAUSE
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYVOLUME) Then lblProfileStep.Caption = "Volumes"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYWC) Then lblProfileStep.Caption = "% Work Cap"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOTARGET) Then lblProfileStep.Caption = "Grams"
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOUNDOLOAD) Then lblProfileStep.Caption = "Grams"
                                lblStepMinutes.Caption = "Volumes"
                                pnlStepActual.Caption = Format(StationControl(Index, Index2).Actual, "###0.0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYVOLUME) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGEBYWC) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOTARGET) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0.0")
                                If (StationRecipe(Index, Index2).Purge_Method = PURGETOUNDOLOAD) Then pnlStepMax.Caption = Format(StationControl(Index, Index2).Target, "###0.0")
                                pnlMinutesActual.Caption = Format(PurgeControl(Index, Index2).Purge_Volumes, "#,##0.0")
                                pnlMinutesMax.Caption = Format(StationRecipe(Index, Index2).Purge_MaxVolumes, "#,##0")
                                Select Case PurgeControl(Index, Index2).curStep
                                    Case 0
                                        ' pauseing
                                        sStr = "PurgePause #"
                                        sStr = sStr & Format(PurgeControl(Index, Index2).curCycle, "##0")
                                        sStr = sStr & " - " & Format((StationRecipe(Index, Index2).Purge_TargetPause - PurgeControl(Index, Index2).StepElapsedMinutes), "###0.00")
                                        sStr = sStr & " min to go"
                                    Case 1
                                        ' purging
                                        sStr = "Purging #"
                                        sStr = sStr & Format(PurgeControl(Index, Index2).curCycle, "##0")
                                        sStr = sStr & " - " & Format((StationRecipe(Index, Index2).Purge_TargetPurge - PurgeControl(Index, Index2).StepElapsedMinutes), "###0.00")
                                        sStr = sStr & " min to go"
                                End Select
                                pnlProfile.Caption = sStr
                        End Select
                        
                    Case PURGEBYTIME
                        ' Purging by Time
                        ' Purge By Time doesn't display first row (stepmax & step)
                        pnlProfile.Top = DelayBoxTop
                        pnlProfile.Caption = "Purge By Time"
                        lblProfileStep.Left = OutOfSight
                        pnlStepActual.Left = OutOfSight
                        pnlStepMax.Left = OutOfSight
                        lblStepMinutes.Left = BoxLblLeft
                        pnlMinutesActual.Left = BoxActLeft
                        pnlMinutesMax.Left = BoxMaxLeft
                        ' update elapsed minutes (and target minutes)
                        lblStepMinutes.Caption = "max/minutes"
                        pnlMinutesActual.Caption = Format((PurgeControl(Index, Index2).ElapsedHours * CSng(60)), "###0.00")
                        pnlMinutesMax.Caption = Format(StationRecipe(Index, Index2).Purge_Time, "###0.00")
                    
                    Case Else
                        pnlProfile.Top = OutOfSight
                        pnlStepActual.Caption = Format(0, "####0")
                        pnlStepMax.Caption = Format(0, "####0")
                        pnlMinutesActual.Caption = Format(0, "####0")
                        pnlMinutesMax.Caption = Format(0, "####0")
                End Select
                
            Else
                pnlProfile.Top = OutOfSight
                pnlStepActual.Caption = Format(0, "####0")
                pnlStepMax.Caption = Format(0, "####0")
                pnlMinutesActual.Caption = Format(0, "####0")
                pnlMinutesMax.Caption = Format(0, "####0")
            End If
                
        Case VBLOAD
            If (LoadControl(Index, Index2).Phase = LoadLoading) Then
    
                Select Case StationRecipe(Index, Index2).Load_Method
                    Case LOADBYTIME
                        ' Loading by Time
                        ' Load By Time doesn't display first row (stepmax & step)
                        pnlProfile.Top = DelayBoxTop
                        pnlProfile.Caption = "Load By Time"
                        lblProfileStep.Left = OutOfSight
                        pnlStepActual.Left = OutOfSight
                        pnlStepMax.Left = OutOfSight
                        lblStepMinutes.Left = BoxLblLeft
                        pnlMinutesActual.Left = BoxActLeft
                        pnlMinutesMax.Left = BoxMaxLeft
                        ' update elapsed minutes (and target minutes)
                        lblStepMinutes.Caption = "max/minutes"
                        pnlMinutesActual.Caption = Format((LoadControl(Index, Index2).ElapsedHours * CSng(60)), "###0.00")
                        pnlMinutesMax.Caption = Format(StationRecipe(Index, Index2).Load_Time, "###0.00")
                    
                    Case Else
                        pnlProfile.Top = OutOfSight
                        pnlStepActual.Caption = Format(0, "####0")
                        pnlStepMax.Caption = Format(0, "####0")
                        pnlMinutesActual.Caption = Format(0, "####0")
                        pnlMinutesMax.Caption = Format(0, "####0")
                End Select
                
            Else
                pnlProfile.Top = OutOfSight
                pnlStepActual.Caption = Format(0, "####0")
                pnlStepMax.Caption = Format(0, "####0")
                pnlMinutesActual.Caption = Format(0, "####0")
                pnlMinutesMax.Caption = Format(0, "####0")
            End If
                
        Case Else
            pnlProfile.Top = OutOfSight
            pnlStepActual.Caption = Format(0, "####0")
            pnlStepMax.Caption = Format(0, "####0")
            pnlMinutesActual.Caption = Format(0, "####0")
            pnlMinutesMax.Caption = Format(0, "####0")
       
    End Select
                
    
    '**********************
    '**********************
    '**********************
    '  Delay Display Panel
    '  Delay Display Panel
    '  Delay Display Panel
    '**********************
    '**********************
    '**********************
    pnlDelay.Left = frmStnDetail.pnlCycle.Left
    Select Case StationControl(Index, Index2).Mode
           
        Case VBPOSTLEAK
            pnlDelay.Top = DelayBoxTop
    '        pnlDelay.Align = 1
            pnlDelay.Caption = "Pause after Leak"
            temptime = StationControl(Index, Index2).End_Time - Now()
            tempSec = CLng((3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime))
            pauseSec = CLng(60# * StationRecipe(Index, Index2).PauseLeakTime)
            pnlToGo.Caption = Format(tempSec, "####0")
            pnlTotal.Caption = Format(pauseSec, "####0")
            pnlToGo.ToolTipText = "Delay Remaining in Seconds"
            pnlTotal.ToolTipText = "Total Delay in Seconds"
        
        Case VBPRELOAD
            pnlDelay.Top = DelayBoxTop
    '        pnlDelay.Align = 1
            pnlDelay.Caption = "PreLoad N2Push"
            temptime = StationControl(Index, Index2).End_Time - Now()
            tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
            pauseSec = CLng(StationConfig(Index, Index2).NitrogenPurgeTime)
            pnlToGo.Caption = Format(tempSec, "####0")
            pnlTotal.Caption = Format(pauseSec, "####0")
            pnlToGo.ToolTipText = "Delay Remaining in Seconds"
            pnlTotal.ToolTipText = "Total Delay in Seconds"
        
        Case VBLOAD
            ' Loading or Waiting for Scales to Settle?
            If LoadControl(Index, Index2).Phase = LoadPause Then
                ' Waiting for Scales to Settle
                pnlDelay.Top = DelayBoxTop
                pnlDelay.Caption = "Load Settling Time"
                temptime = LoadControl(Index, Index2).PhaseDts - Now()
                tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
                pauseSec = CLng(60 * StationConfig(Index, Index2).LoadSettleTime)
                pnlToGo.Caption = Format(tempSec, "####0")
                pnlTotal.Caption = Format(pauseSec, "####0")
                pnlToGo.ToolTipText = "Delay Remaining in Seconds"
                pnlTotal.ToolTipText = "Total Delay in Seconds"
            Else
                ' Loading
                pnlDelay.Top = OutOfSight
                pnlDelay.Caption = "*****"
                pnlToGo.Caption = Format(0, "####0")
                pnlTotal.Caption = Format(0, "####0")
            End If
            
            
        Case VBPOSTLOAD
            pnlDelay.Top = DelayBoxTop
    '        pnlDelay.Align = 1
            pnlDelay.Caption = "Pause after Load"
            temptime = StationControl(Index, Index2).End_Time - Now()
            tempSec = CLng((3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime))
            pauseSec = CLng(60# * StationRecipe(Index, Index2).PauseLoadTime)
            pnlToGo.Caption = Format(tempSec, "####0")
            pnlTotal.Caption = Format(pauseSec, "####0")
            pnlToGo.ToolTipText = "Delay Remaining in Seconds"
            pnlTotal.ToolTipText = "Total Delay in Seconds"
        
        Case VBPURGE
                ' Purging or Waiting for Scales to Settle?
                If PurgeControl(Index, Index2).Phase = PurgePause Then
                    ' Waiting for Scales to Settle
                    pnlDelay.Top = DelayBoxTop
                    pnlDelay.Caption = "Purge Settling Time"
                    temptime = PurgeControl(Index, Index2).PhaseDts - Now()
                    tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
                    pauseSec = CLng(60 * StationConfig(Index, Index2).PurgeSettleTime)
                    pnlToGo.Caption = Format(tempSec, "####0")
                    pnlTotal.Caption = Format(pauseSec, "####0")
                    pnlToGo.ToolTipText = "Delay Remaining in Seconds"
                    pnlTotal.ToolTipText = "Total Delay in Seconds"
                Else
                    ' Purging
                    pnlDelay.Top = OutOfSight
                    pnlDelay.Caption = "*****"
                    pnlToGo.Caption = Format(0, "####0")
                    pnlTotal.Caption = Format(0, "####0")
                End If
            
            
        Case VBPOSTPURGE
            pnlDelay.Top = DelayBoxTop
    '        pnlDelay.Align = 1
            pnlDelay.Visible = True
            pnlDelay.Caption = "Pause after Purge"
            temptime = StationControl(Index, Index2).End_Time - Now()
            tempSec = CLng((3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime))
            pauseSec = CLng(60# * StationRecipe(Index, Index2).PausePurgeTime)
            pnlToGo.Caption = Format(tempSec, "####0")
            pnlTotal.Caption = Format(pauseSec, "####0")
            pnlToGo.ToolTipText = "Delay Remaining in Seconds"
            pnlTotal.ToolTipText = "Total Delay in Seconds"
        
        Case VBSTARTWAIT
            pnlDelay.Top = DelayBoxTop
    '        pnlDelay.Align = 1
            pnlDelay.Caption = "Delayed Start"
            ' TOTAL
            tempHr = 0
            tempMin = 0
            tempSec = 0
            dSeconds = StationControl(Index, Index2).DelaySeconds
            While dSeconds >= 3600
                tempHr = tempHr + 1
                dSeconds = dSeconds - 3600
            Wend
            While dSeconds >= 60
                tempMin = tempMin + 1
                dSeconds = dSeconds - 60
            Wend
            tempSec = dSeconds
            If StationControl(Index, Index2).DelaySeconds > 3600 Then           ' 3600 sec = 1 hour
                ' Display hh:mm:ss
                pnlTotal.Caption = Format(tempHr, "#####0") & ":" & Format(tempMin, "00") & ":" & Format(tempSec, "00")
                pnlTotal.ToolTipText = "Total Delay in hr:min:sec"
            ElseIf StationControl(Index, Index2).DelaySeconds > 60 Then        ' 60 sec = 1 minute
                ' Display mm:ss
                pnlTotal.Caption = Format(tempMin, "00") & ":" & Format(tempSec, "00")
                pnlTotal.ToolTipText = "Total Delay in min:sec"
            Else
                ' Display ss
                pnlTotal.Caption = Format(tempSec, "#0")
                pnlTotal.ToolTipText = "Total Delay in Seconds"
            End If
            ' TO GO
            tempHr = 0
            tempMin = 0
            tempSec = 0
            dSeconds = StationControl(Index, Index2).DelayToGo
            While dSeconds >= 3600
                tempHr = tempHr + 1
                dSeconds = dSeconds - 3600
            Wend
            While dSeconds >= 60
                tempMin = tempMin + 1
                dSeconds = dSeconds - 60
            Wend
            tempSec = dSeconds
            If StationControl(Index, Index2).DelayToGo > 3600 Then           ' 3600 sec = 1 hour
                ' Display hh:mm:ss
                pnlToGo.Caption = Format(tempHr, "#####0") & ":" & Format(tempMin, "00") & ":" & Format(tempSec, "00")
                pnlToGo.ToolTipText = "Delay Remaining in hr:min:sec"
            ElseIf StationControl(Index, Index2).DelayToGo > 60 Then        ' 60 sec = 1 minute
                ' Display mm:ss
                pnlToGo.Caption = Format(tempMin, "00") & ":" & Format(tempSec, "00")
                pnlToGo.ToolTipText = "Delay Remaining in min:sec"
            Else
                ' Display ss
                pnlToGo.Caption = Format(tempSec, "#0")
                pnlToGo.ToolTipText = "Delay Remaining in Seconds"
            End If
        
        Case VBCOURSEPAUSE
            pnlDelay.Top = DelayBoxTop
    '        pnlDelay.Align = 1
            pnlDelay.Visible = True
            pnlDelay.Caption = "Course Pause"
            tempMin = CLng(StationSequence(Index, Index2).CourseData(StationControl(Index, Index2).Course).PauseDuration)
            temptime = StationSequence(Index, Index2).CourseData(StationControl(Index, Index2).Course).DtsStart + TimeSerial(0, CInt(tempMin), 0) - Now()
            tempSec = CLng((3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime))
            pauseSec = CLng(60# * tempMin)
            pnlToGo.Caption = Format(tempSec, "####0")
            pnlTotal.Caption = Format(pauseSec, "####0")
            pnlToGo.ToolTipText = "Delay Remaining in Seconds"
            pnlTotal.ToolTipText = "Total Delay in Seconds"
    
        Case VBCOURSEWAIT
            pnlDelay.Top = DelayBoxTop
    '        pnlDelay.Align = 1
            pnlDelay.Visible = True
            pnlDelay.Caption = "Course Wait for Operator"
            tempMin = DateDiff("n", StationSequence(Index, Index2).CourseData(StationControl(Index, Index2).Course).DtsStart, Now)
            tempSec = DateDiff("s", StationSequence(Index, Index2).CourseData(StationControl(Index, Index2).Course).DtsStart, Now)
            pnlTotal.Caption = Format(tempMin, "##,##0") & ":" & Format((tempSec Mod 60), "00")
            pnlTotal.ToolTipText = "Total Elapsed Wait in Minutes:Seconds"
            lblToGo.Left = OutOfSight
            pnlToGo.Left = OutOfSight
    
        Case Else
    '        pnlDelay.Align = 0
            pnlDelay.Top = OutOfSight
            pnlDelay.Caption = "*****"
            lblToGo.Left = lblTotal.Left
            pnlToGo.Left = pnlTotal.Left
            pnlToGo.Caption = Format(0, "####0")
            pnlTotal.Caption = Format(0, "####0")
       
    End Select
    
    '**********************
    '**********************
    '**********************
    ' Leakcheck Display
    ' Leakcheck Display
    ' Leakcheck Display
    '**********************
    '**********************
    '**********************
    If StationControl(Index, Index2).Mode = VBLEAK Then
        pnlLeakcheck.Top = DelayBoxTop
        pnlLeakcheck.Caption = "Leakcheck"
        pnlLcStepDesc(0).Caption = LeakPhaseDesc(LeakCheckControl.Phase) & " " & LeakMethodDesc(LeakCheckControl.Method)
        Select Case LeakCheckControl.Phase
            Case LeakPurging
                pnlLcStepDesc(1).Caption = "ToGo: " & DurationDescription(CSng(DateDiff("s", Now, LeakCheckControl.PhaseDts)) / CSng(60))
            Case LeakPressurizing
                pnlLcStepDesc(1).Caption = "PressToGo: " & Format((StationConfig(Index, Index2).LCSetPoint - PTinvalue), "###0.0") & " psi"
            Case LeakTesting
                pnlLcStepDesc(1).Caption = "ToGo: " & DurationDescription(CSng(DateDiff("s", Now, LeakCheckControl.PhaseDts)) / CSng(60))
            Case LeakComplete
                pnlLcStepDesc(1).Caption = "Complete"
            Case Else
                pnlLcStepDesc(1).Caption = "future"
        End Select
    Else
        pnlLeakcheck.Top = OutOfSight
    End If
    
    '**********************
    '**********************
    '**********************
    ' LC Press Display
    ' LC Press Display
    ' LC Press Display
    '**********************
    '**********************
    '**********************
    If StationControl(Index, Index2).Mode = VBLEAK Then
        pnlLcPressBox.Top = PurgeDpBoxTop
        pnlLcPress.ForeColor = TitlesData_Forecolor
        pnlLcPress.Caption = Format(Com_AIO(acComnPressSensor).EUValue, "##0.00")
    Else
        pnlLcPressBox.Top = OutOfSight
        pnlLcPress.ForeColor = DKGRAY
        pnlLcPress.Caption = "0.00"
    End If
    
    '**********************
    '**********************
    '**********************
    ' Course Display
    ' Course Display
    ' Course Display
    '**********************
    '**********************
    '**********************
    If (StationSequence(Index, Index2).NumCourses > 1) Then
        pnlCourse.Top = CourseBoxTop
        If (StationControl(Index, Index2).Mode = VBIDLE) Then
            pnlCourse.ForeColor = DKGRAY
            pnlCourse.Caption = "Course " & StationControl(Index, Index2).Course & " of " & _
                StationSequence(Index, Index2).NumCourses
        Else
            pnlCourse.ForeColor = TitlesData_Forecolor
            pnlCourse.Caption = "Course " & (StationControl(Index, Index2).Course) & " of " & _
                StationSequence(Index, Index2).NumCourses
        End If
    Else
        pnlCourse.Top = OutOfSight
    End If
    
    '**********************
    '**********************
    '**********************
    ' Cycle Display
    ' Cycle Display
    ' Cycle Display
    '**********************
    '**********************
    '**********************
    If StationControl(Index, Index2).Mode = VBIDLE Then
        pnlCycle.Top = OutOfSight
        pnlCycle.ForeColor = DKGRAY
        Select Case StationRecipe(Index, Index2).EndMethod
            Case ENDCYCLES
                pnlCycle.Caption = "Cycle " & StationControl(Index, Index2).CompletedCycles & " of " & _
                    StationRecipe(Index, Index2).Cycles
            Case ENDWEIGHTCHG
                pnlCycle.Caption = "Cycle " & StationControl(Index, Index2).CompletedCycles
            Case Else
                pnlCycle.Caption = "Cycle " & StationControl(Index, Index2).CompletedCycles & " of " & _
                    StationRecipe(Index, Index2).Cycles
        End Select
    Else
        pnlCycle.Top = CycleBoxTop
        pnlCycle.ForeColor = TitlesData_Forecolor
        Select Case StationRecipe(Index, Index2).EndMethod
            Case ENDCYCLES
                pnlCycle.Caption = "Cycle " & StationControl(Index, Index2).CurrCycle & " of " & _
                    StationRecipe(Index, Index2).Cycles
            Case ENDWEIGHTCHG
                pnlCycle.Caption = "Cycle " & StationControl(Index, Index2).CurrCycle
            Case Else
                pnlCycle.Caption = "Cycle " & StationControl(Index, Index2).CurrCycle & " of " & _
                    StationRecipe(Index, Index2).Cycles
        End Select
    End If
    
    '**********************
    '**********************
    '**********************
    ' PurgeDP Display
    ' PurgeDP Display
    ' PurgeDP Display
    '**********************
    '**********************
    '**********************
    If (USINGPURGEDP) And StationControl(Index, Index2).Mode = VBPURGE Then
        pnlPurgeDpBox.Top = PurgeDpBoxTop
        pnlPurgeDp.ForeColor = TitlesData_Forecolor
        pnlPurgeDp.Caption = Format(Stn_AIO(Index, asPurgeDiffPress).EUValue, "##0.00")
    Else
        pnlPurgeDpBox.Top = OutOfSight
        pnlPurgeDp.ForeColor = DKGRAY
        pnlPurgeDp.Caption = "0.00"
    End If
    
    '**********************
    '**********************
    '**********************
    ' WaterBath Display
    ' WaterBath Display
    ' WaterBath Display
    '**********************
    '**********************
    '**********************
    If (USINGWATERBATH And STN_INFO(Index).ADF_DEF.hasADF_WaterBath And (StationControl(Index, Index2).Mode <> VBIDLE)) Then
        pnlWaterBath.Top = WaterBathBoxTop
        Select Case StationConfig(Index, Index2).WaterBathControl
            Case wbDirect
                pnlWaterBath.Caption = "Direct WaterBath Control"
            Case wbFuelTemp
                pnlWaterBath.Caption = "WaterBath Controls Fuel Temp"
            Case wbVaporTemp
                pnlWaterBath.Caption = "WaterBath Control Vapor Temp"
        End Select
        pnlSuperSp.Caption = Format(PID_INFO(wbSuperTemp).SP, "###0.0##")
        pnlSuperPv.Caption = Format(PID_INFO(wbSuperTemp).PV, "###0.0##")
        pnlSuperPv.ForeColor = TitlesData_Forecolor
        pnlWbSp.Caption = Format(LF_Chiller.SpIn, "###0.0##")
        pnlWbPv.Caption = Format(LF_Chiller.PvIn, "###0.0##")
        pnlWbPv.ForeColor = TitlesData_Forecolor
    Else
        pnlWaterBath.Top = OutOfSight
    End If
    
    '**********************
    '**********************
    '**********************
    ' "Use TC's" Button
    ' "Use TC's" Button
    ' "Use TC's" Button
    '**********************
    '**********************
    '**********************
    If USINGSTNTC Then
        If Stn_UseTC(Index, Index2) Then
            pnlTC.Top = pnlPurgeDpBox.Top + 375 + 15
            pnlTC1.ForeColor = TitlesData_Forecolor
            pnlTC2.ForeColor = TitlesData_Forecolor
            pnlTC1.Caption = Format(Stn_AIO(Index, asStationTC1).EUValue, "##0.0")
            pnlTC2.Caption = Format(Stn_AIO(Index, asStationTC2).EUValue, "##0.0")
        Else
            pnlTC.Top = OutOfSight
        End If
        cmdUseTC.Top = pnlPurgeDpBox.Top + 375 + 15 + 60
        cmdUseTC.Caption = IIf(Stn_UseTC(Index, Index2), "ON", "OFF")
    Else
        pnlTC.Top = OutOfSight
        cmdUseTC.Top = OutOfSight
    End If
    
    '**********************
    '**********************
    '**********************
    ' "Use Thermo" Button
    ' "Use Thermo" Button
    ' "Use Thermo" Button
    '**********************
    '**********************
    '**********************
    If USINGCOMMONTC Then
'        cmdThermo.Top = tabsLegend.Top - (cmdThermo.Height + 15)  ' was 4540
        cmdThermo.Top = 4850
    Else
        cmdThermo.Top = OutOfSight
    End If
    
    
    ' **************************************************************************************
    ' ************************************  OVERSIZE BOXES *********************************
    ' ************************************  OVERSIZE BOXES *********************************
    ' ************************************  OVERSIZE BOXES *********************************
    ' **************************************************************************************

    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    '  Report File Name Text Entry
    '  Report File Name Text Entry
    '  Report File Name Text Entry
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    If StationControl(Index, Index2).Mode = VBIDLE Then
        If SysConfig.ReportFileName1stPart = RPT_OPERENTER _
          Or SysConfig.ReportFileName2ndPart = RPT_OPERENTER _
          Or SysConfig.ReportFileName3rdPart = RPT_OPERENTER Then
            With txtRptMsg
                .FontSize = 10
                .Top = 1225
                .Height = 735
                .Left = 120
                .Width = pbxRptName.Width - 240
            End With
            pbxRptName.Top = OversizeBoxTop
            pbxRptName.Left = OversizeBoxLeft
            If SysConfig.ReportFileName1stPart = RPT_OPERENTER Then
                txtRptName1.Visible = True
            Else
                txtRptName1.Visible = False
            End If
            If SysConfig.ReportFileName2ndPart = RPT_OPERENTER Then
                txtRptName2.Visible = True
                txtRptName2.Left = txtRptName1.Left
            Else
                txtRptName2.Visible = False
            End If
            If SysConfig.ReportFileName3rdPart = RPT_OPERENTER Then
                txtRptName3.Visible = True
                txtRptName3.Left = txtRptName1.Left
            Else
                txtRptName3.Visible = False
            End If
        Else
            pbxRptName.Top = OutOfSight
        End If
    Else
        pbxRptName.Top = OutOfSight
    End If
    cmdApproved.Picture = IIf(Stn_OperReportNameIsValid, cmdApproved_OK.Picture, cmdApproved_No.Picture)
    cmdApproved.ToolTipText = IIf(Stn_OperReportNameIsValid, cmdApproved_OK.ToolTipText, cmdApproved_No.ToolTipText)
    
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    '  Station Sequence Display Panel
    '  Station Sequence Display Panel
    '  Station Sequence Display Panel
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ChgErrModule 89, 1311
    If SEQ_Nmbr(Index, Index2) <> seqIdle Then
        pnlStnSeq.Top = OversizeBoxTop
        pnlStnSeq.Caption = SEQ_Task(Index, Index2)
        txtSeqMsg.Visible = True
        txtSeqMsg.text = SEQ_Message(Index, Index2)
    Else
        pnlStnSeq.Top = OutOfSight
        txtSeqMsg.Visible = False
    End If
    
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    '  CANVENTALARM Override Display Panel
    '  CANVENTALARM Override Display Panel
    '  CANVENTALARM Override Display Panel
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    If USINGCANVENTALARM And SysConfig.CanVent_Delay_Max > 0 Then
        pnlCanVentOvr.Top = OversizeBoxTop
        If OOTs(Index, Index2).CanVent_DelayOn = True Then
            pnlCanVentOvr.BackColor = Common_BackColor
            pnlCanVentOvr.ForeColor = MEDRED
            pnlCanVentOvr.Caption = "        Can Vent Flow Switch Override is Active       (" + Format(OOTs(Index, Index2).CanVent_DelayCount, "####0")
            pnlCanVentOvr.Caption = pnlCanVentOvr.Caption + " of " + Format(SysConfig.CanVent_Delay_Max, "####0") + " sec)"
        Else
            pnlCanVentOvr.BackColor = Common_BackColor
            pnlCanVentOvr.ForeColor = Black
            pnlCanVentOvr.Caption = "Can Vent Flow Switch Override is Off"
        End If
    Else
        pnlCanVentOvr.Top = OutOfSight
    End If
    
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ' LOAD RATE UPDATE DISPLAY
    ' LOAD RATE UPDATE DISPLAY
    ' LOAD RATE UPDATE DISPLAY
    '       only display if user = APS
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ChgErrModule 89, 1330
    If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And (StationRecipe(Index, Index2).LiveFuel)) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(Index, Index2).LiveFuel)) Or (Not CheckPass("H", False))) Then
        pnlLoadRate.Top = OutOfSight
    Else
        If StationControl(Index, Index2).Mode = VBLOAD _
          And StationRecipe(Index, Index2).Load_Method = LOADBYWEIGHT _
          And LoadControl(Index, Index2).Phase = LoadLoading _
          And StationRecipe(Index, Index2).Load_Wt = StationCanister(Index, Index2).WorkingCapacity Then
            pnlLoadRate.Top = OversizeBoxTop
        Else
            pnlLoadRate.Top = OutOfSight
        End If
    End If
    
    
    
    
    ' **************************************************************************************
    ' ************************************  SPAWNED BOXES **********************************
    ' ************************************  SPAWNED BOXES **********************************
    ' ************************************  SPAWNED BOXES **********************************
    ' **************************************************************************************

    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ' Net Weight Change Display
    ' Net Weight Change Display
    ' Net Weight Change Display
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ChgErrModule 89, 1330
    If (StationRecipe(Index, Index2).EndMethod = ENDWEIGHTCHG) Then
        flagShowIt = True
        ' don't show if station is idle
        If (StationControl(Index, Index2).Mode = VBIDLE) Then flagShowIt = False
        ' don't show if station was idle before AllStationPause
        If ((StationControl(Index, Index2).Mode = VBPAUSEALARM) And (StationControl(Index, Index2).Mode_PauseSave = VBIDLE)) Then flagShowIt = False
        ' Open Net Weight Change Screen
        If flagShowIt Then
            ' if it is not open already
            If (Not LoadControl(Index, Index2).NetWtChgIsOpen) Then frmNetWtChg.Show
        End If
    End If
    
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ' LOAD CONCORDANCE DISPLAY
    ' LOAD CONCORDANCE DISPLAY
    ' LOAD CONCORDANCE DISPLAY
    '       only display if user = APS
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ChgErrModule 89, 1331
    If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And (StationRecipe(Index, Index2).LiveFuel)) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(Index, Index2).LiveFuel)) Or (Not CheckPass("G", False))) Then
        ' no concordance
    Else
        If StationControl(Index, Index2).Mode = VBLOAD _
          And LoadControl(Index, Index2).Phase = LoadLoading _
          And StationRecipe(Index, Index2).UsePriScale = True Then
            If StationControl(Index, Index2).TestTimer > Stn_LoadEql_StartTimer(Index, Index2) Then
                ' Open Concordance Screen
                If (Not LoadControl(Index, Index2).ConcordanceIsOpen) Then frmConcordance.Show
            End If
        End If
    End If
        
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ' update StnDetail variables
    ' update StnDetail variables
    ' update StnDetail variables
    ' **************************************************************************************
    ' **************************************************************************************
    ' **************************************************************************************
    ChgErrModule 89, 1349
    stnDtl_DispStn_Last = Index
    stnDtl_DispShift_Last = Index2
    stnDtl_StnCourse_Last(Index, Index2) = StationControl(Index, Index2).Course
    stnDtl_StnMode_Last(Index, Index2) = StationControl(Index, Index2).Mode
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Private Sub Update_Text(Index As Integer, Index2 As Integer)
Dim sTxt As String

    'Update Text fields on detail screen
    pnlStnName.Caption = STN_INFO(Index).desc
    sTxt = StationSequence(Index, Index2).EstSeqDurDesc
    txtRecipeName.text = StationRecipe(Index, Index2).Name
    txtRcpDsc(0).text = StationRecipe(Index, Index2).desc(0)
    txtRcpDsc(1).text = StationRecipe(Index, Index2).desc(1)
    txtRcpDsc(2).text = StationRecipe(Index, Index2).desc(2)
    txtCanID.text = StationCanister(Index, Index2).Description
    Select Case StationControl(Index, Index2).LeakCheckStatus
        Case RESULTGOOD
            txtLeakCheckStatus.ForeColor = Good_ForeColor
            txtLeakCheckStatus.text = "Passed LeakCheck"
'            txtLeakCheckStatus.text = StationControl(Index, index2).LcStatusDescription & " LeakCheck"
        Case NORESULT
            txtLeakCheckStatus.ForeColor = txtLeakCheckStatus.BackColor
'            txtLeakCheckStatus.ForeColor = MEDGRAY
            txtLeakCheckStatus.text = StationControl(Index, Index2).LcStatusDescription
        Case Else
            txtLeakCheckStatus.ForeColor = Alarm_ForeColor
            txtLeakCheckStatus.text = "Failed LeakCheck"
'            txtLeakCheckStatus.text = " LeakCheck " & StationControl(Index, index2).LcStatusDescription
    End Select
    lblBedVolume.Caption = Format(StationCanister(Index, Index2).WorkingVolume, "###0.00")
    lblWorkCap.Caption = Format(StationCanister(Index, Index2).WorkingCapacity, "###0.00")
    lblWorkCap.ForeColor = IIf((StationCanister(Index, Index2).WorkingCapacity = CSng(0)), Warning_ForeColor, Black)
    txtEngineer.text = JobInfo(Index, Index2).Engineer
    txtVehicle.text = JobInfo(Index, Index2).Vehicle
    txtStartOp.text = JobInfo(Index, Index2).Start_Op
    txtEndOp.text = JobInfo(Index, Index2).End_Op
'    txtComment.text = JobInfo(Index, index2).Comment
    txtRptMsg.ForeColor = IIf(Stn_OperReportNameIsValid, Message_ForeColor, MEDRED)
'    Set txtRptMsg.MultiLine = False
    txtRptMsg.Width = IIf(Stn_OperReportNameIsValid, txtStnDtlMsg.Width + 300, txtStnDtlMsg.Width)
    txtRptMsg.text = IIf(Stn_OperReportNameIsValid, (vbCrLf & "Report Filename Text Accepted"), txtRptMsg.text)
    txtRptName1.BackColor = IIf(Stn_OperReportNameIsValid, txtEngineer.BackColor, txtRptName1.BackColor)
    txtRptName2.BackColor = IIf(Stn_OperReportNameIsValid, txtEngineer.BackColor, txtRptName2.BackColor)
    txtRptName3.BackColor = IIf(Stn_OperReportNameIsValid, txtEngineer.BackColor, txtRptName3.BackColor)
    Select Case StationControl(Index, Index2).Mode
        Case VBIDLE, VBIDLEWAITING
            txtStnDtlMsg.ForeColor = Message_ForeColor
        Case VBCOURSEWAIT
            txtStnDtlMsg.text = vbCrLf & StationSequence(Index, Index2).CourseData(StationControl(Index, Index2).Course).MsgText
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBCOURSEPAUSE
            sTxt = DurationDescription(StationSequence(Index, Index2).CourseData(StationControl(Index, Index2).Course).PauseDuration)
            txtStnDtlMsg.text = vbCrLf & Trim(StationControl(Index, Index2).Job_Description) & " Paused for " & sTxt
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBPURGEWAIT
            If (USINGPURGEOVEN And StationRecipe(Index, Index2).PurgeOven And (Not PurgeControl(Index, Index2).PurgeOvenTempOK)) Then
                txtStnDtlMsg.text = vbCrLf & "waiting for Purge Oven"
            Else
                txtStnDtlMsg.text = vbCrLf & "waiting"
            End If
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBSCALEWAIT, VBSHIFTWAIT, VBSTARTWAIT, VBLEAKWAIT, VBPURGEWAIT
            txtStnDtlMsg.text = vbCrLf & "waiting"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBPAUSE, VBFIDPAUSE
            txtStnDtlMsg.text = vbCrLf & "paused"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBPAUSEBYUSER
            txtStnDtlMsg.text = vbCrLf & "Paused by User"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBGASPAUSE
            txtStnDtlMsg.text = vbCrLf & "Paused for Vapor Tank Refill"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBWBPAUSE
            txtStnDtlMsg.text = vbCrLf & "Waiting for WaterBath Temp"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBPAUSEOOT
            txtStnDtlMsg.text = vbCrLf & "Paused due to an OOT"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBPAUSEALARM
            txtStnDtlMsg.text = vbCrLf & "Paused due to an Alarm"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case VBLEAKERROR
            txtStnDtlMsg.text = vbCrLf & "Leakcheck Error"
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
        Case Else
            sTxt = Trim(StationControl(Index, Index2).Job_Description)
            txtStnDtlMsg.text = vbCrLf & "Running " & sTxt
            txtStnDtlMsg.ForeColor = ModeBackColor(StationControl(Index, Index2).Mode)
    End Select
End Sub

Private Sub ShiftDown()
Dim lstShift
    lstShift = DispShift
    DispShift = IIf(DispShift <= 1, NR_SHIFT, DispShift - 1)
    If DispShift <> lstShift Then
        txtStnDtlMsg.text = " "
        stnDtl_StnMode_Last(DispStn, DispShift) = -1        ' force update of station mode indicator
        Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
        ChartXYValues DispStn, DispShift
        Update_Text DispStn, DispShift
        Update_Stn DispStn, DispShift
        If frmRecipe.Visible Then frmRecipe.RecipeDisplay_ByStnShift               ' force update of station recipe
    End If
End Sub

Private Sub ShiftUp()
Dim lstShift
    lstShift = DispShift
    DispShift = IIf(DispShift = NR_SHIFT, 1, DispShift + 1)
    If DispShift <> lstShift Then
        txtStnDtlMsg.text = " "
        stnDtl_StnMode_Last(DispStn, DispShift) = -1        ' force update of station mode indicator
        Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
        ChartXYValues DispStn, DispShift
        Update_Text DispStn, DispShift
        Update_Stn DispStn, DispShift
        If frmRecipe.Visible Then frmRecipe.RecipeDisplay_ByStnShift               ' force update of station recipe
    End If
End Sub

Private Sub AlarmLog()
    If StationControl(DispStn, DispShift).DBFile <> "" Then
       View_Alarm StationControl(DispStn, DispShift).Job_Number, DispStn, DispShift
    Else
       txtStnDtlMsg.text = "Station MUST be running to view STATION ALARMS"
    End If
End Sub

Public Sub StationContinue(ByVal byWhom As String)
    If Pause_Alarm Then
        txtStnDtlMsg.text = "Can not CONTINUE while system is paused"
        Exit Sub                                    ' Wait for no alarming conditions
    End If
    If CheckPass("R", msgSHOW) Then
        '   reset (optional) Local PAS Timeout Timers
        If USINGPASLOCALCONTROL Then
            If PAS_INFO(pasTEMPERATURE).timeOut Then
                PAS_INFO(pasTEMPERATURE).TimeOutDuration = 0#
                PAS_INFO(pasTEMPERATURE).timeOut = False
            End If
            If PAS_INFO(pasMOISTURE).timeOut Then
                PAS_INFO(pasMOISTURE).TimeOutDuration = 0#
                PAS_INFO(pasMOISTURE).timeOut = False
            End If
        End If
        ' Set the Continue-Button-Has-Been-Pressed flag for this Station
        StationControl(DispStn, DispShift).ContinueRequest = True
        ' build msg for Stn Detail screen
        If InStr(tbrStnDetail.Buttons("start").ToolTipText, "Reset") > 0 Then
            sStr = Right(tbrStnDetail.Buttons("start").ToolTipText, Len(tbrStnDetail.Buttons("start").ToolTipText) - 6)
            sStr = sStr & " reset by " & byWhom
        ElseIf InStr(tbrStnDetail.Buttons("start").ToolTipText, "Resume") > 0 Then
            sStr = Right(tbrStnDetail.Buttons("start").ToolTipText, Len(tbrStnDetail.Buttons("start").ToolTipText) - 7)
            sStr = sStr & " resumed by " & byWhom
        ElseIf InStr(tbrStnDetail.Buttons("start").ToolTipText, "Cancel") > 0 Then
            sStr = Right(tbrStnDetail.Buttons("start").ToolTipText, Len(tbrStnDetail.Buttons("start").ToolTipText) - 7)
            sStr = sStr & " cancelled by " & byWhom
        Else
            sStr = "Operator continued " & ModeDescShort(StationControl(DispStn, DispShift).Mode)
        End If
        ' display msg
        txtStnDtlMsg.text = sStr
    End If
End Sub

Public Sub StationPause(ByVal byWhom As String)
'
'  Pause a station because Operator pressed Pause
'
      
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 89, 3762
      
    ' start pause time
    StationControl(DispStn, DispShift).PausedDts = Now
    Write_ELog byWhom & " Paused Station #" & Format(DispStn, "0") & " Shift #" & Format(DispShift, "0")
    Write_JLog DispStn, DispShift, byWhom & " Paused Station"
    
    ' Save the current mode for the continue button
    StationControl(DispStn, DispShift).Mode_PauseSave = StationControl(DispStn, DispShift).Mode
    ' save elapsed hours so far
    Select Case StationControl(DispStn, DispShift).Mode
        Case VBLEAK
            LeakCheckControl.ElapsedHours_Prev = LeakCheckControl.ElapsedHours
        Case VBLOAD
            LoadControl(DispStn, DispShift).ElapsedHours_Prev = LoadControl(DispStn, DispShift).ElapsedHours
        Case VBPURGE
            PurgeControl(DispStn, DispShift).ElapsedHours_Prev = PurgeControl(DispStn, DispShift).ElapsedHours
        Case Else
    End Select
   
    If StationControl(DispStn, DispShift).Mode_PauseSave = VBLOAD Then              ' station was loading before Pause
        If StationRecipe(DispStn, DispShift).Load_Method = LOADBYTIME Or StationRecipe(DispStn, DispShift).Load_Method = LOADBYWC Then
            StationControl(DispStn, DispShift).PauseAlarmStartTime = Now            ' save pause time on load by time
        End If
    End If
   
    If StationControl(DispStn, DispShift).Mode_PauseSave = VBPURGE Then             ' station was purging before Pause
        StationControl(DispStn, DispShift).PauseAlarmStartTime = Now                ' save pause time
    End If
    
   
    '  Turn Off Station MFCs
    ShutdownStnMFCs DispStn, DispShift
    '   Station Valves
    Close_Stn_Valves DispStn, DispShift
    '   Scale Valves
    If StationRecipe(DispStn, DispShift).UsePriScale And StationControl(DispStn, DispShift).PriScaleStn > 0 _
            And StationControl(DispStn, DispShift).PriScaleStn < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(DispStn, DispShift).PriScaleStn, isPriAuxVentSol, cOFF
    End If
    If StationRecipe(DispStn, DispShift).PurgeAuxCan And StationControl(DispStn, DispShift).AuxScaleStn > 0 Then
        Stn_OutDigital StationControl(DispStn, DispShift).AuxScaleStn, isAuxPurgeSol, cOFF
    End If
    ' Release Common (Leak) Pressure Transducer (if this station is using it)
    If LeakCheckControl.station = DispStn Then
        LeakCheckControl.station = 0
        LeakCheckControl.Shift = 0
        LeakCheckControl.Phase = 0
        LeakCheckControl.ElapsedHours = 0
        LeakCheckControl.ElapsedHours_Prev = 0
    End If
    
    ' set mode to Paused
    Select Case byWhom
        Case "Oper", "Operator"
            StationControl(DispStn, DispShift).Mode = VBPAUSEBYUSER
        Case "ADF"
            StationControl(DispStn, DispShift).Mode = VBGASPAUSE
        Case Else
            StationControl(DispStn, DispShift).Mode = VBPAUSEBYUSER
    End Select
          
 ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Private Sub StnDown()
Dim lstStn
    lstStn = DispStn
    DispStn = IIf(DispStn <= 1, LAST_STN, DispStn - 1)
    If DispStn <> lstStn Then
        If (STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Then
            frmLeakTest.Show
        Else
            txtStnDtlMsg.text = " "
            stnDtl_StnMode_Last(DispStn, DispShift) = -1
            Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
            ChartXYValues DispStn, DispShift
            txtLoadRateSp.text = Format(StationRecipe(DispStn, DispShift).Load_Rate, "###0.00")
            Update_Text DispStn, DispShift
            Update_Stn DispStn, DispShift
            If frmRecipe.Visible Then frmRecipe.RecipeDisplay_ByStnShift               ' force update of station recipe
        End If
    End If
End Sub

Private Sub StnUp()
Dim lstStn
    lstStn = DispStn
    DispStn = IIf(DispStn >= LAST_STN, 1, DispStn + 1)
    If DispStn <> lstStn Then
        If (STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Then
            frmLeakTest.Show
        Else
            txtStnDtlMsg.text = " "
            stnDtl_StnMode_Last(DispStn, DispShift) = -1
            Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
            ChartXYValues DispStn, DispShift
            txtLoadRateSp.text = Format(StationRecipe(DispStn, DispShift).Load_Rate, "###0.00")
            Update_Text DispStn, DispShift
            Update_Stn DispStn, DispShift
            If frmRecipe.Visible Then frmRecipe.RecipeDisplay_ByStnShift               ' force update of station recipe
        End If
    End If
End Sub

Private Sub cmdDnShift_Click()
    ShiftDown
End Sub

Private Sub cmdDnStn_Click()
    StnDown
End Sub

Private Sub cmdLoadRateUpdate_Click()
Dim minGramPerHour, maxGramPerHour, newGramPerHour, maxSLPM, reqSLPM As Single
Dim sGramsPerLiter As Single

    ' MAKE SURE THAT THE NEW LOAD RATE IS VALID
    If IsEmpty(txtLoadRateSp.text) Then txtLoadRateSp.text = Format(StationRecipe(DispStn, DispShift).Load_Rate, "###0.00")
    If Not IsNumeric(txtLoadRateSp.text) Then txtLoadRateSp.text = Format(StationRecipe(DispStn, DispShift).Load_Rate, "###0.00")
    Select Case STN_INFO(DispStn).Type
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            ' Butane is mixed with the N2
            sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
            minGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
            maxGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
            If CSng(txtLoadRateSp.text) < minGramPerHour Then txtLoadRateSp.text = Format(minGramPerHour, "###0.00")
            If CSng(txtLoadRateSp.text) > maxGramPerHour Then txtLoadRateSp.text = Format(maxGramPerHour, "###0.00")
            ' now is the mix % greater than the capabilities of the flow controller
            reqSLPM = ((100 - StationRecipe(DispStn, DispShift).Mix_Percent) / StationRecipe(DispStn, DispShift).Mix_Percent) * GramsPerHourToSlpm(CSng(txtLoadRateSp.text), sGramsPerLiter)
            maxSLPM = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
            If reqSLPM > maxSLPM Then
                ' fix it
                newGramPerHour = (maxSLPM * StationRecipe(DispStn, DispShift).Mix_Percent) / (100 - StationRecipe(DispStn, DispShift).Mix_Percent)
                txtLoadRateSp.text = Format(newGramPerHour, "###0.00")
            End If
        Case STN_ORVR2_TYPE
            ' Butane is mixed with the N2
            If StationRecipe(DispStn, DispShift).UseHiRangeMFC Then
                ' use higher range MFC
                sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfc2DensityMult
                minGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                maxGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.95)
            Else
                ' use lower range MFC
                sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                minGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                maxGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
            End If
            If CSng(txtLoadRateSp.text) < minGramPerHour Then txtLoadRateSp.text = Format(minGramPerHour, "###0.00")
            If CSng(txtLoadRateSp.text) > maxGramPerHour Then txtLoadRateSp.text = Format(maxGramPerHour, "###0.00")
            ' now is the mix % greater than the capabilities of the flow controller
            If StationRecipe(DispStn, DispShift).UseHiRangeMFC Then
                ' use higher range MFC
                sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfc2DensityMult
                reqSLPM = ((100 - StationRecipe(DispStn, DispShift).Mix_Percent) / StationRecipe(DispStn, DispShift).Mix_Percent) * GramsPerHourToSlpm(CSng(txtLoadRateSp.text), sGramsPerLiter)
                maxSLPM = 0.95 * Stn_AIO(DispStn, asNitrogenORVRFlow).EuMax
            Else
                ' use lower range MFC
                sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                reqSLPM = ((100 - StationRecipe(DispStn, DispShift).Mix_Percent) / StationRecipe(DispStn, DispShift).Mix_Percent) * GramsPerHourToSlpm(CSng(txtLoadRateSp.text), sGramsPerLiter)
                maxSLPM = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
            End If
            If reqSLPM > maxSLPM Then
                ' fix it
                newGramPerHour = (maxSLPM * StationRecipe(DispStn, DispShift).Mix_Percent) / (100 - StationRecipe(DispStn, DispShift).Mix_Percent)
                txtLoadRateSp.text = Format(newGramPerHour, "###0.00")
            End If
        Case STN_LIVEFUEL_TYPE
            ' LiveFuel Vapor is carried by the Nitrogen
            ' nothing to do - This screen is not enabled if this Station is a LiveFuel station
        Case STN_LIVEREG_TYPE
            If (StationRecipe(DispStn, DispShift).LiveFuel) Then
                ' LiveFuel Vapor is carried by the Nitrogen
                ' nothing to do - This screen is not enabled if this Station is a LiveFuel station
            Else
                ' Butane is mixed with the N2
                If StationRecipe(DispStn, DispShift).UseHiRangeMFC Then
                    ' use higher range MFC
                    sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfc2DensityMult
                    minGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                    maxGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.95)
                Else
                    ' use lower range MFC
                    sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                    minGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                    maxGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
                End If
                If CSng(txtLoadRateSp.text) < minGramPerHour Then txtLoadRateSp.text = Format(minGramPerHour, "###0.00")
                If CSng(txtLoadRateSp.text) > maxGramPerHour Then txtLoadRateSp.text = Format(maxGramPerHour, "###0.00")
                ' now is the mix % greater than the capabilities of the flow controller
                If StationRecipe(DispStn, DispShift).UseHiRangeMFC Then
                    ' use higher range MFC
                    sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfc2DensityMult
                    reqSLPM = ((100 - StationRecipe(DispStn, DispShift).Mix_Percent) / StationRecipe(DispStn, DispShift).Mix_Percent) * GramsPerHourToSlpm(CSng(txtLoadRateSp.text), sGramsPerLiter)
                    maxSLPM = 0.95 * Stn_AIO(DispStn, asNitrogenORVRFlow).EuMax
                Else
                    ' use lower range MFC
                    sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                    reqSLPM = ((100 - StationRecipe(DispStn, DispShift).Mix_Percent) / StationRecipe(DispStn, DispShift).Mix_Percent) * GramsPerHourToSlpm(CSng(txtLoadRateSp.text), sGramsPerLiter)
                    maxSLPM = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
                End If
                If reqSLPM > maxSLPM Then
                    ' fix it
                    newGramPerHour = (maxSLPM * StationRecipe(DispStn, DispShift).Mix_Percent) / (100 - StationRecipe(DispStn, DispShift).Mix_Percent)
                    txtLoadRateSp.text = Format(newGramPerHour, "###0.00")
                End If
            End If
        Case STN_LIVEORVR2_TYPE
            If (StationRecipe(DispStn, DispShift).LiveFuel) Then
                ' LiveFuel Vapor is carried by the Nitrogen
                ' nothing to do - This screen is not enabled if this Station is a LiveFuel station
            Else
                ' Butane is mixed with the N2
                sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                minGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                maxGramPerHour = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
                If CSng(txtLoadRateSp.text) < minGramPerHour Then txtLoadRateSp.text = Format(minGramPerHour, "###0.00")
                If CSng(txtLoadRateSp.text) > maxGramPerHour Then txtLoadRateSp.text = Format(maxGramPerHour, "###0.00")
                ' now is the mix % greater than the capabilities of the flow controller
                reqSLPM = ((100 - StationRecipe(DispStn, DispShift).Mix_Percent) / StationRecipe(DispStn, DispShift).Mix_Percent) * GramsPerHourToSlpm(CSng(txtLoadRateSp.text), sGramsPerLiter)
                maxSLPM = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
                If reqSLPM > maxSLPM Then
                    ' fix it
                    newGramPerHour = (maxSLPM * StationRecipe(DispStn, DispShift).Mix_Percent) / (100 - StationRecipe(DispStn, DispShift).Mix_Percent)
                    txtLoadRateSp.text = Format(newGramPerHour, "###0.00")
                End If
            End If
        Case STN_COMBO3_TYPE
            ' future
        Case Else
            ' nothing to do
    End Select
            
    ' IMPLEMENT THE NEW LOAD RATE
    StationRecipe(DispStn, DispShift).Load_Rate = CSng(txtLoadRateSp.text)
    LoadSetPoint_Update DispStn, DispShift
    
End Sub

Private Sub StationStart()
    If CheckPass("R", msgSHOW) Then
    
        If Pause_Alarm <> 0 Then                                ' error, in alarm
            txtStnDtlMsg.text = "Can not START while system is paused"
            Exit Sub
        End If
        If StationControl(DispStn, DispShift).Mode <> VBIDLE Then          ' error, in use
            txtStnDtlMsg.text = "Start Button pushed when not in IDLE...Error"
            Exit Sub
        End If
        
        ' if setup failed, notify user
        If Not StationSequence(DispStn, DispShift).Validated Then
            txtStnDtlMsg.text = "Invalid Job Sequence; Nothing to do"
            Exit Sub
        End If
        
        If Stn_OperReportNameIsValid = False _
            And _
                (SysConfig.ReportFileName1stPart = RPT_OPERENTER _
                Or SysConfig.ReportFileName2ndPart = RPT_OPERENTER _
                Or SysConfig.ReportFileName3rdPart = RPT_OPERENTER) Then
            txtStnDtlMsg.text = "Must have Validated Report File Name before Starting"
            Exit Sub
        End If
    
        
        ' set "Start Station" flag (used by Module 2)
        StationControl(DispStn, DispShift).StartRequest = True
        txtStnDtlMsg.text = vbCrLf & "Starting Station #" & Format(DispStn, "#0")
        Write_ELog "Starting Station #" & Format(DispStn, "#0")
        
    End If

End Sub

Private Sub StationStop()
    If CheckPass("R", msgSHOW) Then
        frmStop.Show
        frmStop.cmdYES = False
    End If
End Sub

Private Sub StatisticsSummary()
    frmSummary.Show
End Sub

Private Sub cmdThermo_Click()
    frmCommonTC.Show
End Sub

Private Sub OOTLog()
    If StationControl(DispStn, DispShift).DBFile <> "" Then
        View_OOT StationControl(DispStn, DispShift).Job_Number, DispStn, DispShift
    Else
        txtStnDtlMsg.text = "Station MUST be running for Tolerances"
    End If
End Sub

Private Sub JobLog()
    If StationControl(DispStn, DispShift).DBFile <> "" Then
        View_JobLog StationControl(DispStn, DispShift).Job_Number, DispStn, DispShift
    Else
        txtStnDtlMsg.text = "Station MUST be running for Job Event Log"
    End If
End Sub

Private Sub cmdUpShift_Click()
    ShiftUp
End Sub

Private Sub cmdUpStn_Click()
    StnUp
End Sub

Private Sub cmdUseTC_Click()
    If USINGSTNTC Then
        If CheckPass("U", True) Then
            Stn_UseTC(DispStn, DispShift) = Not Stn_UseTC(DispStn, DispShift)
            If Stn_UseTC(DispStn, DispShift) Then
                cmdUseTC.Caption = "TCs: ON"
            Else
                cmdUseTC.Caption = "TCs: OFF"
            End If
        End If
    End If
End Sub

Private Sub ValidateFilename()
Dim validText As Boolean
Dim sChar, sMsg As String
Dim Idx As Integer

    ' Anything to check?
    If Not SysConfig.ReportFileName1stPart = RPT_OPERENTER _
        And Not SysConfig.ReportFileName2ndPart = RPT_OPERENTER _
        And Not SysConfig.ReportFileName3rdPart = RPT_OPERENTER Then
            Stn_OperReportNameIsValid = True
            Exit Sub
    End If
    
    ' check Filename
    sMsg = ""
    validText = True
    If SysConfig.ReportFileName1stPart = RPT_OPERENTER Then
        txtRptName1.text = Trim(txtRptName1.text)
        If Len(txtRptName1.text) = 0 Then
            txtRptName1.BackColor = txtEngineer.BackColor
        Else
            For Idx = 1 To Len(txtRptName1.text)
                sChar = Mid(txtRptName1.text, Idx, 1)
                Select Case Asc(sChar)
                    Case 32, 33, 35 To 41, 43 To 45, 91, 93 To 96, 123, 125
                        ' valid punctuation characters
                    Case 48 To 57
                        ' 0 - 9
                    Case 65 To 90
                        ' A - Z
                    Case 97 To 122
                        ' a - z
                    Case Else
                        ' invalid
                        If Not validText Then sMsg = sMsg & vbCrLf
                        sMsg = sMsg & ">>> " & sChar & " <<< Not Allowed in Report Name"
                        txtRptMsg.text = sMsg
                        txtRptName1.BackColor = EntryInvalid_BackColor
                        validText = False
                End Select
            Next Idx
        End If
    Else
        txtRptName1.text = ""
    End If
        
    If SysConfig.ReportFileName2ndPart = RPT_OPERENTER Then
        txtRptName2.text = Trim(txtRptName2.text)
        If Len(txtRptName2.text) = 0 Then
            txtRptName2.BackColor = txtEngineer.BackColor
        Else
            For Idx = 1 To Len(txtRptName2.text)
                sChar = Mid(txtRptName2.text, Idx, 1)
                Select Case Asc(sChar)
                    Case 32, 33, 35 To 41, 43 To 45, 91, 93 To 96, 123, 125
                        ' valid punctuation characters
                    Case 48 To 57
                        ' 0 - 9
                    Case 65 To 90
                        ' A - Z
                    Case 97 To 122
                        ' a - z
                    Case Else
                        ' invalid
                        If Not validText Then sMsg = sMsg & vbCrLf
                        sMsg = sMsg & ">>> " & sChar & " <<< Not Allowed in Report Name"
                        txtRptMsg.text = sMsg
                        txtRptName1.BackColor = EntryInvalid_BackColor
                        validText = False
                End Select
            Next Idx
        End If
    Else
        txtRptName2.text = ""
    End If
        
    If SysConfig.ReportFileName3rdPart = RPT_OPERENTER Then
        txtRptName3.text = Trim(txtRptName3.text)
        If Len(txtRptName3.text) = 0 Then
            txtRptName3.BackColor = txtEngineer.BackColor
        Else
            For Idx = 1 To Len(txtRptName3.text)
                sChar = Mid(txtRptName3.text, Idx, 1)
                Select Case Asc(sChar)
                    Case 32, 33, 35 To 41, 43 To 45, 91, 93 To 96, 123, 125
                        ' valid punctuation characters
                    Case 48 To 57
                        ' 0 - 9
                    Case 65 To 90
                        ' A - Z
                    Case 97 To 122
                        ' a - z
                    Case Else
                        ' invalid
                        If Not validText Then sMsg = sMsg & vbCrLf
                        sMsg = sMsg & ">>> " & sChar & " <<< Not Allowed in Report Name"
                        txtRptMsg.text = sMsg
                        txtRptName1.BackColor = EntryInvalid_BackColor
                        validText = False
                End Select
            Next Idx
        End If
    Else
        txtRptName3.text = ""
    End If
        
    ' if flag is false then text is invalid
    Stn_OperReportNameIsValid = validText

End Sub

Private Sub cmdValidate_Click()
    ValidateFilename
    Update_Text DispStn, DispShift
End Sub

Private Sub Form_Activate()
Dim cntr As Integer
    cntr = 0
    Do Until ((STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Or (cntr > NR_STN))
        If (STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Then
            DispStn = IIf((DispStn < NR_STN), (DispStn + 1), 1)
        End If
        cntr = cntr + 1
    Loop
    txtStnDtlMsg.text = " "
    stnDtl_StnMode_Last(DispStn, DispShift) = -1
    Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
    ChartXYValues DispStn, DispShift
    txtLoadRateSp.text = Format(StationRecipe(DispStn, DispShift).Load_Rate, "###0.00")
    Update_Text DispStn, DispShift
    Update_Stn DispStn, DispShift
    UpdateNavigateBtns
End Sub

Private Sub Form_GotFocus()
    UpdateNavigateBtns
    Update_Text DispStn, DispShift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmStnDetail = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Load()
Dim CVleft, PVleft, Upper, Wide As Integer
Dim clr As Long
Dim i, j As Integer

    KeyPreview = True
    frmStnDetail.Height = frmMainMenu.Height
    frmStnDetail.Width = frmMainMenu.Width
    ScreenDebug = False
    
    BuildToolbars
    
'    pbxVertical.Left = 6955
    
    txtStnDtlMsg.text = " "
    txtRptMsg.text = " "
    
    ' Set Foreground colors
    frmStnCanister.ForeColor = Titles_ForeColor
    txtCanID.ForeColor = TitlesData_Forecolor
    frmStnRecipe.ForeColor = Titles_ForeColor
    txtRecipeName.ForeColor = TitlesData_Forecolor
'    frmComment.ForeColor = Titles_ForeColor
    txtDspStn.ForeColor = TitlesData_Forecolor
    txtStation.ForeColor = TitlesData_Forecolor
    pnlNameFrame.ForeColor = TitlesData_Forecolor
    pnlTestTime.ForeColor = TitlesData_Forecolor
    pnlJobDuration.ForeColor = TitlesData_Forecolor
    pnlProfile.ForeColor = TitlesLabel_ForeColor
    pnlMinutesActual.ForeColor = TitlesData_Forecolor
    pnlMinutesMax.ForeColor = ModeBackColor(VBPURGE)
    pnlStepActual.ForeColor = TitlesData_Forecolor
    pnlStepMax.ForeColor = ModeBackColor(VBPURGE)
    pnlDelay.ForeColor = TitlesLabel_ForeColor
    pnlToGo.ForeColor = TitlesData_Forecolor
    pnlTotal.ForeColor = TitlesData_Forecolor
    pnlScale.ForeColor = TitlesLabel_ForeColor
    pnlWtPri.ForeColor = TitlesData_Forecolor
    pnlChgPri.ForeColor = TitlesData_Forecolor
    pnlWtAux.ForeColor = TitlesData_Forecolor
    pnlChgAux.ForeColor = TitlesData_Forecolor
    pnlStnSeq.ForeColor = TitlesLabel_ForeColor
    pnlLeakcheck.ForeColor = TitlesLabel_ForeColor
    pnlLcStepDesc(0).ForeColor = Data_ForeColor ' was ModeBackColor(VBLEAK)
    pnlLcStepDesc(1).ForeColor = Data_ForeColor ' was ModeBackColor(VBLEAK)
    pnlCourse.ForeColor = TitlesData_Forecolor
    pnlCycle.ForeColor = TitlesData_Forecolor
    txtStnDtlMsg.ForeColor = Message_ForeColor
'    txtComment.ForeColor = TitlesData_Forecolor
    txtEndOp.ForeColor = TitlesData_Forecolor
    txtEngineer.ForeColor = TitlesData_Forecolor
    txtStartOp.ForeColor = TitlesData_Forecolor
    txtVehicle.ForeColor = TitlesData_Forecolor
    
    ' report filename boxes
    If SysConfig.ReportFileName1stPart = RPT_OPERENTER _
      Or SysConfig.ReportFileName2ndPart = RPT_OPERENTER _
      Or SysConfig.ReportFileName3rdPart = RPT_OPERENTER Then
        With pbxRptName
            .Width = frmStnDtlMsg.Width
            .Left = frmStnDtlMsg.Left
            .Top = tbrStnDetail.Top + tbrStnDetail.Height
        End With
        If SysConfig.ReportFileName1stPart = RPT_OPERENTER Then
            txtRptName1.Visible = True
        Else
            txtRptName1.Visible = False
        End If
        If SysConfig.ReportFileName2ndPart = RPT_OPERENTER Then
            txtRptName2.Visible = True
        Else
            txtRptName2.Visible = False
        End If
        If SysConfig.ReportFileName3rdPart = RPT_OPERENTER Then
            txtRptName3.Visible = True
        Else
            txtRptName3.Visible = False
        End If
    Else
        pbxRptName.Top = OutOfSight
    End If
    
    BoxLblLeft = 0
    BoxMaxLeft = 1320
    BoxActLeft = 2070
    
    pnlStatus.Left = 90
    pnlStatus.Width = pnlStatusFrame.Width - 210
    pnlReportFrame.Width = frmStnDetail.Width - pnlReportFrame.Left - 135
    pnlReport.Left = 90
    pnlReport.Width = pnlReportFrame.Width - 210
    pnlStnName.Left = 90
    pnlStnName.Width = pnlNameFrame.Width - 210
    pnlStnName.FontSize = 12
    pnlReport.FontSize = 12
    pnlStatus.FontSize = 12
    
'   **********   Status Bar Setup
    pnlAlarms.Left = 0
    pnlAlarms.Width = pnlAlarms.Width - pnlMix.Width
    pnlAlarms.Top = 0
    pnlAlarms.Height = pnlEstop.Height + 150
    
    pnlMessageFrame.Left = pnlAlarms.Left + pnlAlarms.Width
    pnlMessageFrame.Top = pnlAlarms.Top
    pnlMessageFrame.Height = pnlAlarms.Height
    pnlMessage.Left = 60
    pnlMessage.Top = 60
    pnlMessage.Height = pnlMessageFrame.Height - 120
    pnlMessage.Width = pnlMessageFrame.Width - 120
    
    pnlPurgeAir.Left = pnlMessageFrame.Left + pnlMessageFrame.Width
    pnlPurgeAir.Width = frmStnDetail.Width - pnlPurgeAir.Left - 150
    pnlPurgeAir.Top = pnlAlarms.Top
    pnlPurgeAir.Height = pnlAlarms.Height
    
    ' Status Bar Update
    UpdateStatusBars
    
    ' GRAPHS
    WhichGraph = ShowBarGraphs
    ShowXYLegend = False
    tabsWhichGraph.Left = 15
    tabsWhichGraph.Top = 5305
    tabsWhichGraph.Width = pnlCycle.Width
    tabsWhichGraph.TabFixedWidth = 1415
    Set tabsWhichGraph.SelectedItem = tabsWhichGraph.Tabs(2)
    tabsLegend.Left = OutOfSight
    tabsLegend.Top = tabsWhichGraph.Top - tabsWhichGraph.Height - 45
    tabsLegend.Width = tabsWhichGraph.Width
    tabsLegend.TabFixedWidth = tabsWhichGraph.TabFixedWidth
    Set tabsLegend.SelectedItem = tabsLegend.Tabs(2)
    GraphsTop = IIf((pbxRptName.Top = OutOfSight), (pbxBottom.Top - pnlBarGraphs.Height), OutOfSight)
    ' XY Graphs
    pnlXYGraphs.Top = GraphsTop
    pnlXYGraphs.Left = 0
    pnlXYGraphs.Width = pnlBarGraphs.Width
    chtStnChart.Title.text = "NET GRAMS"
    With chtStnChart.Title.VtFont
       .Name = "Arial"
       .Style = VtFontStyleBold
    '       .Effect = VtFontEffectUnderline
       .Size = 12
        clr = DK2GREEN
       .VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
    End With
    chtStnChart.chartType = VtChChartType2dXY  ' set to X Y Scatter chart
    chtStnChart.Plot.UniformAxis = False
    With chtStnChart.Plot.Axis(VtChAxisIdY).ValueScale
        .Auto = False
        .Maximum = 10
        .Minimum = -2
        .MajorDivision = 6
        .MinorDivision = 2
    End With
    chtStnChart.Visible = True
    For i = 1 To LAST_STN
        For j = 1 To NR_SHIFT
            ClearXYvalues i, j
            DataCollector i, j
        Next j
    Next i
    Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
    ChartXYValues DispStn, DispShift
    
    ' Bar Graphs
    pnlBarGraphs.Top = GraphsTop
    pnlBarGraphs.Left = 0
    CVleft = 75
    PVleft = 715
    Upper = 4950
    Wide = 650
    With txtBtnCV
        .Left = CVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlBtnCV.BackColor
        .ForeColor = pnlBtnCV.FloodColor
    End With
    With txtBtnPV
        .Left = PVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlBtnPV.BackColor
        .ForeColor = pnlNitPV.FloodColor
    End With
    With txtNitCV
        .Left = CVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlNitCV.BackColor
        .ForeColor = pnlNitCV.FloodColor
    End With
    With txtNitPV
        .Left = PVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlNitPV.BackColor
        .ForeColor = pnlNitPV.FloodColor
    End With
    With txtPurCV
        .Left = CVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlPurCV.BackColor
        .ForeColor = pnlPurCV.FloodColor
    End With
    With txtPurPV
        .Left = PVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlPurPV.BackColor
        .ForeColor = pnlPurPV.FloodColor
    End With
    With txtTarget
        .Left = CVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlTarget.BackColor
        .ForeColor = pnlTarget.FloodColor
    End With
    With txtActual
        .Left = PVleft
        .Top = Upper
        .Width = Wide
        .Alignment = vbCenter
        .BorderStyle = vbBSNone
        .Enabled = True
        .Locked = True
        .BackColor = pnlActual.BackColor
        .ForeColor = pnlActual.FloodColor
    End With
    
    ' position for "Oversize" Boxes such as LoadRate & CanVent
    OversizeBoxTop = 4280   ' was 8160
    OversizeBoxLeft = 6815  ' was 10200
    OversizeBoxWidth = 5040
    
    ' positions for Big(Height=1185) and Slim(Height=375) Boxes
    EstJobDurBoxTop = 0
    ElapsedBoxTop = EstJobDurBoxTop + 375 + 15
    DelayBoxTop = ElapsedBoxTop + 375 + 15
    CourseBoxTop = DelayBoxTop + 1185 + 15
    CycleBoxTop = CourseBoxTop + 375 + 15
    ScaleBoxTop = CycleBoxTop + 375 + 15
    PurgeDpBoxTop = ScaleBoxTop + 1185 + 15
    WaterBathBoxTop = ScaleBoxTop + 1185 + 15
    
    
    ' boxes on the vertical panel
    ' boxes on the vertical panel
    ' boxes on the vertical panel
    '
    '       SLIM BOXES
    ' cycles panel
    pnlCycle.Left = 0
    pnlCycle.Width = pbxVertical.Width
    pnlCycle.Height = 375
    ' courses panel
    pnlCourse.Left = pnlCycle.Left
    pnlCourse.Width = pnlCycle.Width
    pnlCourse.Height = pnlCycle.Height
    ' elapsed time panel
    pnlTestTime.Left = pnlCycle.Left
    pnlTestTime.Width = pnlCycle.Width
    pnlTestTime.Height = pnlCycle.Height
    ' estimated job duration panel
    pnlJobDuration.Left = pnlCycle.Left
    pnlJobDuration.Width = pnlCycle.Width
    pnlJobDuration.Height = pnlCycle.Height
    ' purgeDP panel
    pnlPurgeDpBox.Left = pnlCycle.Left
    pnlPurgeDpBox.Width = pnlCycle.Width
    pnlPurgeDpBox.Height = pnlCycle.Height
    '
    '       ODD BALL SLIM BOX
    ' display Common TC's panel
    cmdThermo.Left = pnlCycle.Left
    cmdThermo.Width = pnlCycle.Width
    cmdThermo.Height = pnlCycle.Height
    '
    '       BIG BOXES
    ' scale(s) panel
    pnlScale.Left = 0
    pnlScale.Width = pbxVertical.Width
    pnlScale.Height = 1185
        lblScaleDesc.Top = 890
        lblScaleWeight.Top = lblScaleDesc.Top
        lblScaleChange.Top = lblScaleDesc.Top
    ' delay panel
    pnlDelay.Left = pnlScale.Left
    pnlDelay.Width = pnlScale.Width
    pnlDelay.Height = pnlScale.Height
    ' purge profile panel
    pnlProfile.Left = pnlScale.Left
    pnlProfile.Width = pnlScale.Width
    pnlProfile.Height = pnlScale.Height
    ' leakcheck panel
    pnlLeakcheck.Left = pnlScale.Left
    pnlLeakcheck.Width = pnlScale.Width
    pnlLeakcheck.Height = pnlScale.Height
    ' waterbath panel
    pnlWaterBath.Left = pnlScale.Left
    pnlWaterBath.Width = pnlScale.Width
    pnlWaterBath.Height = pnlScale.Height
    
    
    ' OVERSIZE BOXES
    ' OVERSIZE BOXES
    ' OVERSIZE BOXES
    '
    ' change load rate panel
    pnlLoadRate.Left = OversizeBoxLeft
    pnlLoadRate.Width = OversizeBoxWidth
    ' can vent override active message panel
    pnlCanVentOvr.Left = OversizeBoxLeft
    pnlCanVentOvr.Width = OversizeBoxWidth
    ' station sequence panel
    pnlStnSeq.Left = OversizeBoxLeft
    pnlStnSeq.Width = OversizeBoxWidth
    ' Report Name panel
    pbxRptName.Left = OversizeBoxLeft
    pbxRptName.Width = OversizeBoxWidth
    
    
    ' display "current" station and shift
    txtDspStn.text = DispStn
    txtDspShift.text = DispShift
    
    ' update station detail screen
    stnDtl_StnMode_Last(DispStn, DispShift) = -1        ' force update of station mode indicator
    Update_Stn DispStn, DispShift
    Update_Text DispStn, DispShift
    
End Sub

Private Sub CloseScreen()
    Unload Me
    Set frmStnDetail = Nothing
End Sub

Private Sub mnuAirLog_Click()
    menuViewAirLog
End Sub

Private Sub mnuCopyFile_Click()
    menuCopyFile
End Sub

Private Sub mnuCourses_Click()
    menuCourses
End Sub

Private Sub mnuFirstAid_Click()
    menuFirstAid
End Sub

Private Sub mnuFuelUseLog_Click()
    menuViewFuelUseLog
End Sub

Private Sub mnuOotMonitor_Click()
    menuOotMonitor
End Sub

Private Sub mnuPrintFile_Click()
    menuPrintFile
End Sub

Private Sub mnuOperatorManual_Click()
    menuOperatorManual
End Sub

Private Sub mnuTomCanLoad_Click()
    menuRemCanLoad
End Sub

Private Sub pbxTop_DblClick()
    If CheckPass("H", False) Then
        ScreenDebug = IIf(ScreenDebug, False, True)
    End If
End Sub

Private Sub tabsLegend_Click()
    Select Case tabsLegend.SelectedItem
        Case "Legend"
            ShowXYLegend = True
            chtStnChart.ShowLegend = True
        Case "No Legend"
            ShowXYLegend = False
            chtStnChart.ShowLegend = False
    End Select
End Sub

Private Sub tabsWhichGraph_Click()
    Select Case tabsWhichGraph.SelectedItem
        Case "MFC Flow"
            WhichGraph = ShowBarGraphs
            tabsLegend.Left = OutOfSight
        Case "Net Grams"
            WhichGraph = ShowXYGraphs
            tabsLegend.Left = tabsWhichGraph.Left
            Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
    End Select
End Sub

Private Sub tmrScreen_Timer()
    If DispStn > 0 And DispShift > 0 Then
        If (OptoReadAllOnce Or Not IoComOn) Then Update_Stn DispStn, DispShift
    Else
        DispStn = 1
        DispShift = 1
    End If
End Sub

Private Sub tmrXYGraphs_Timer()
Dim iStn, iShift As Integer
    For iStn = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
            ' See if time to update XY Chart data
            If (Stn_XYGraph_TestTimer(iStn, iShift) + Stn_XYChart_Xinterval(iStn, iShift)) <= StationControl(iStn, iShift).TestTimer Then
                ' reset XYGraph timer
                Stn_XYGraph_TestTimer(iStn, iShift) = Stn_XYGraph_TestTimer(iStn, iShift) + Stn_XYChart_Xinterval(iStn, iShift)     ' # of seconds
                ' update XYGraph Data only if not paused for anything
                If Not StationControl(iStn, iShift).IsPausedInAlarm Then
                    ' update data
                    DataCollector iStn, iShift
                    ' update chart
                    If (iStn = DispStn) And (iShift = DispShift) Then ChartXYValues iStn, iShift
                End If
            End If
        Next iShift
    Next iStn
End Sub

Private Sub txtEndOp_GotFocus()
    txtEndOp.SelStart = 0
    txtEndOp.SelLength = Len(txtEndOp.text)
End Sub

Private Sub txtEndOp_KeyPress(KeyAscii As Integer)
    If CheckPass("Q", True) Then
     If KeyAscii = vbKeyReturn Then
       JobInfo(DispStn, DispShift).End_Op = txtEndOp.text
     End If
    Else
      KeyAscii = 0
    End If
End Sub

Private Sub txtEndOp_LostFocus()
    JobInfo(DispStn, DispShift).End_Op = txtEndOp.text
End Sub

Private Sub txtEngineer_GotFocus()
    txtEngineer.SelStart = 0
    txtEngineer.SelLength = Len(txtEngineer.text)
End Sub

Private Sub txtEngineer_KeyPress(KeyAscii As Integer)
    If CheckPass("Q", True) Then
      If KeyAscii = vbKeyReturn Then
        JobInfo(DispStn, DispShift).Engineer = txtEngineer.text
      End If
    Else
     KeyAscii = 0
    End If
End Sub

Private Sub txtEngineer_LostFocus()
    JobInfo(DispStn, DispShift).Engineer = txtEngineer.text
End Sub

Private Sub txtRptName1_Change()
    txtRptName1.BackColor = txtEngineer.BackColor
    Stn_OperReportNameIsValid = False
    txtRptMsg.text = " "
End Sub

Private Sub txtRptName2_Change()
    txtRptName2.BackColor = txtEngineer.BackColor
    Stn_OperReportNameIsValid = False
    txtRptMsg.text = " "
End Sub

Private Sub txtRptName3_Change()
    txtRptName3.BackColor = txtEngineer.BackColor
    Stn_OperReportNameIsValid = False
    txtRptMsg.text = " "
End Sub

Private Sub txtStartOp_GotFocus()
    txtStartOp.SelStart = 0
    txtStartOp.SelLength = Len(txtStartOp.text)
End Sub

Private Sub txtStartOp_KeyPress(KeyAscii As Integer)
    If CheckPass("Q", True) Then
      If KeyAscii = vbKeyReturn Then
        JobInfo(DispStn, DispShift).Start_Op = txtStartOp.text
      End If
    Else
      KeyAscii = 0
    End If
End Sub

Private Sub txtStartOp_LostFocus()
    JobInfo(DispStn, DispShift).Start_Op = txtStartOp.text
End Sub

Private Sub txtVehicle_GotFocus()
    txtVehicle.SelStart = 0
    txtVehicle.SelLength = Len(txtVehicle.text)
End Sub

Private Sub txtVehicle_KeyPress(KeyAscii As Integer)
    If CheckPass("Q", True) Then
      If KeyAscii = vbKeyReturn Then
        JobInfo(DispStn, DispShift).Vehicle = txtVehicle.text
      End If
    Else
      KeyAscii = 0
    End If
End Sub

Private Sub txtVehicle_LostFocus()
    JobInfo(DispStn, DispShift).Vehicle = txtVehicle.text
End Sub

Private Sub BuildToolbars()
' Create object variable for the Toolbar.
Dim btnX As MSComctlLib.Button
    
    ' ******************
    ' NAVIGATION TOOLBAR
    ' ******************
    
    ' Load the ImageLists
    tbrNavigate.ImageList = frmMainMenu.imgNavigateNormal
    tbrNavigate.DisabledImageList = frmMainMenu.imgNavigateDisabled
    tbrNavigate.HotImageList = frmMainMenu.imgNavigateHot
    
    ' Add button objects to Buttons collection using the
    ' Add method. After creating each button, set both
    ' Description and ToolTipText properties.
    
'    tbrNavigate.Buttons.Add , , , tbrSeparator
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
'    'Login Screen
'    Set btnX = tbrNavigate.Buttons.Add(, "login", , tbrDefault, "login")
'    btnX.ToolTipText = "User Login"
'    btnX.Description = btnX.ToolTipText
'    'Logout
'    Set btnX = tbrNavigate.Buttons.Add(, "logout", , tbrDefault, "logout")
'    btnX.ToolTipText = "User Logout"
'    btnX.Description = btnX.ToolTipText
'
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'
'    'Copy Files
'    Set btnX = tbrNavigate.Buttons.Add(, "copyfiles", , tbrDefault, "copyfiles")
'    btnX.ToolTipText = "Copy Files"
'    btnX.Description = btnX.ToolTipText
'    'Print Files
'    Set btnX = tbrNavigate.Buttons.Add(, "printfiles", , tbrDefault, "printfiles")
'    btnX.ToolTipText = "Print Files"
'    btnX.Description = btnX.ToolTipText

'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)

    'Canisters Screen
    Set btnX = tbrNavigate.Buttons.Add(, "canisters", , tbrDefault, "can_master")
    btnX.ToolTipText = "Master Canisters"
    btnX.Description = btnX.ToolTipText
    'Recipes Screen
    Set btnX = tbrNavigate.Buttons.Add(, "recipes", , tbrDefault, "rcp_master")
    btnX.ToolTipText = "Master Recipes"
    btnX.Description = btnX.ToolTipText
    'Purge Profiles Screen
    Set btnX = tbrNavigate.Buttons.Add(, "purgeprofile", , tbrDefault, "prof_master")
    btnX.ToolTipText = "Master Purge Profiles"
    btnX.Description = btnX.ToolTipText
    'Sequence (Courses) Screen
    Set btnX = tbrNavigate.Buttons.Add(, "courses", , tbrDefault, "seq_master")
    btnX.ToolTipText = "Master Sequences"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    'TOM Can Load Tasks Screen
    Set btnX = tbrNavigate.Buttons.Add(, "tomcanload", , tbrDefault, "remotecontrol")
    btnX.ToolTipText = "Task Order Manager Tasks"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    'Configuration Screen
    Set btnX = tbrNavigate.Buttons.Add(, "configuration", , tbrDefault, "configuration")
    btnX.ToolTipText = "Configuration"
    btnX.Description = btnX.ToolTipText
    'System Definition Screen
    Set btnX = tbrNavigate.Buttons.Add(, "sysdef", , tbrDefault, "sysdef")
    btnX.ToolTipText = "System Definition Screen"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    'Fuel Use Log
    Set btnX = tbrNavigate.Buttons.Add(, "fueluselog", , tbrDefault, "fueluselog")
    btnX.ToolTipText = "Fuel Consumption Log"
    btnX.Description = btnX.ToolTipText
    'Butane Available
    Set btnX = tbrNavigate.Buttons.Add(, "butane", , tbrDefault, "flammablegas")
    btnX.ToolTipText = "Butane Available"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
        
    'Event Log Screen
    Set btnX = tbrNavigate.Buttons.Add(, "eventlog", , tbrDefault, "eventlog")
    btnX.ToolTipText = "Event Log"
    btnX.Description = btnX.ToolTipText
    'Joblist Screen
    Set btnX = tbrNavigate.Buttons.Add(, "joblist", , tbrDefault, "joblist")
    btnX.ToolTipText = "List of Jobs"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
     
    'Station Detail Screen
    Set btnX = tbrNavigate.Buttons.Add(, "stndetail", , tbrDefault, "stndetail")
    btnX.ToolTipText = "Station Detail"
    btnX.Description = btnX.ToolTipText
    'Overview Screen
    Set btnX = tbrNavigate.Buttons.Add(, "overview", , tbrDefault, "overview")
    btnX.ToolTipText = "Overview"
    btnX.Description = btnX.ToolTipText
    'Review Screen
    Set btnX = tbrNavigate.Buttons.Add(, "reviewdata", , tbrDefault, "reviewdata")
    btnX.ToolTipText = "Review Data"
    btnX.Description = btnX.ToolTipText
    'Watch Screen
    Set btnX = tbrNavigate.Buttons.Add(, "watchdata", , tbrDefault, "watchdata")
    btnX.ToolTipText = "Watch Data"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
       
    'Calibration Screen
    Set btnX = tbrNavigate.Buttons.Add(, "calibration", , tbrDefault, "calibration")
    btnX.ToolTipText = "Calibration"
    btnX.Description = btnX.ToolTipText
    'I/O Monitor Screen
    Set btnX = tbrNavigate.Buttons.Add(, "iomonitor", , tbrDefault, "iomonitor")
    btnX.ToolTipText = "I/O Monitor"
    btnX.Description = btnX.ToolTipText
    'Scale Monitor Screen
    Set btnX = tbrNavigate.Buttons.Add(, "scalemonitor", , tbrDefault, "scalemonitor")
    btnX.ToolTipText = "Scale Monitor"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
       
    'Simulation Control Panel
    Set btnX = tbrNavigate.Buttons.Add(, "simulation", , tbrDefault, "simulation")
    btnX.ToolTipText = "Simulation Control Panel"
    btnX.Description = btnX.ToolTipText
    
    If ((Com_DIO(icAlarmBeacon).addr <> 0) Or (Com_DIO(icAlarmBeacon).chan <> 0)) Then
        'TurnOff Beacon
        Set btnX = tbrNavigate.Buttons.Add(, "beaconoff", , tbrDefault, "beaconoff")
        btnX.ToolTipText = "Turn Off Beacon"
        btnX.Description = btnX.ToolTipText
    End If
    
    If ((Com_DIO(icAlarmHorn).addr <> 0) Or (Com_DIO(icAlarmHorn).chan <> 0)) Then
        'TurnOff Horn
        Set btnX = tbrNavigate.Buttons.Add(, "hornoff", , tbrDefault, "hornoff")
        btnX.ToolTipText = "Silence Horn"
        btnX.Description = btnX.ToolTipText
    End If
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            
    'Operators Manual
    Set btnX = tbrNavigate.Buttons.Add(, "opermanual", , tbrDefault, "opermanual")
    btnX.ToolTipText = "Operators Manual"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
                            
'    'FirstAid
'    Set btnX = tbrNavigate.Buttons.Add(, "firstaid", , tbrDefault, "firstaid")
'    btnX.ToolTipText = "FirstAid File Save"
'    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            
    ' blank space
'    Set btnX = tbrNavigate.Buttons.Add(, "fillright", , tbrPlaceholder)
'    btnX.Width = 2550 ' Placeholder width
    
    'Close Screen
'    Set btnX = tbrNavigate.Buttons.Add(, "close", , tbrDefault, "close")
'    btnX.ToolTipText = "Close Screen"
'    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    

    
    
    ' **********************
    ' STATION DETAIL TOOLBAR
    ' **********************
    
    ' Load the ImageLists
    tbrStnDetail.ImageList = imgStnDetailNormal
    tbrStnDetail.DisabledImageList = imgStnDetailDisabled
    tbrStnDetail.HotImageList = imgStnDetailHot
    
'    tbrStnDetail.Buttons.Add , , , tbrSeparator
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    
    ' blank space
'    Set btnX = tbrStnDetail.Buttons.Add(, "fillleft", , tbrPlaceholder)
'    btnX.Width = 1500 ' Placeholder width to accommodate a textbox.
                
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
                
    'Alarm Log
    Set btnX = tbrStnDetail.Buttons.Add(, "alarmlog", , tbrDefault, "alarmlog")
    btnX.ToolTipText = "Alarm Log"
    btnX.Description = btnX.ToolTipText
    'OOT Log
    Set btnX = tbrStnDetail.Buttons.Add(, "ootlog", , tbrDefault, "ootlog")
    btnX.ToolTipText = "Out Of Tolerance Log"
    btnX.Description = btnX.ToolTipText
    'Statistics Summary
    Set btnX = tbrStnDetail.Buttons.Add(, "statsum", , tbrDefault, "statsum")
    btnX.ToolTipText = "Statistics Summary"
    btnX.Description = btnX.ToolTipText
    'Job Log
    Set btnX = tbrStnDetail.Buttons.Add(, "joblog", , tbrDefault, "joblog")
    btnX.ToolTipText = "Job Log"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
                
    'Operator Comment
    Set btnX = tbrStnDetail.Buttons.Add(, "opercomment", , tbrDefault, "opercomment")
    btnX.ToolTipText = "Operator Comment"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    'Start
    Set btnX = tbrStnDetail.Buttons.Add(, "start", , tbrDefault, "start")
    btnX.ToolTipText = "Start Job"
    btnX.Description = btnX.ToolTipText
    'Continue
    Set btnX = tbrStnDetail.Buttons.Add(, "continue", , tbrDefault, "continue")
    btnX.ToolTipText = "Continue Job"
    btnX.Description = btnX.ToolTipText
    btnX.Visible = False
    'Pause
    Set btnX = tbrStnDetail.Buttons.Add(, "pause", , tbrDefault, "pause")
    btnX.ToolTipText = "Pause Job"
    btnX.Description = btnX.ToolTipText
    btnX.Visible = False
    
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    
    'Stop
    Set btnX = tbrStnDetail.Buttons.Add(, "stop", , tbrDefault, "stop")
    btnX.ToolTipText = "Stop Job"
    btnX.Description = btnX.ToolTipText
                
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
                
'    'Station Number
'    Set btnX = tbrStnDetail.Buttons.Add(, "prevstn", , tbrDefault, "prev")
'    btnX.ToolTipText = "Previous Station"
'    btnX.Description = btnX.ToolTipText
'    Set btnX = tbrStnDetail.Buttons.Add(, "StnNoTxt", , tbrPlaceholder)
'    btnX.Width = 750 ' Placeholder width to accommodate a textbox.
'    Set btnX = tbrStnDetail.Buttons.Add(, "nextstn", , tbrDefault, "next")
'    btnX.ToolTipText = "Next Station"
'    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
                
'    'Shift Number
'    Set btnX = tbrStnDetail.Buttons.Add(, "prevshift", , tbrDefault, "prev")
'    btnX.ToolTipText = "Previous Shift"
'    btnX.Description = btnX.ToolTipText
'    Set btnX = tbrStnDetail.Buttons.Add(, "ShiftNoTxt", , tbrPlaceholder)
'    btnX.Width = 500 ' Placeholder width to accommodate a textbox.
'    Set btnX = tbrStnDetail.Buttons.Add(, "nextshift", , tbrDefault, "next")
'    btnX.ToolTipText = "Next Shift"
'    btnX.Description = btnX.ToolTipText
    
    'Station Canister
    Set btnX = tbrStnDetail.Buttons.Add(, "canisters", , tbrDefault, "canisters")
    btnX.ToolTipText = "Station Canister"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    
    'Station Recipe
    Set btnX = tbrStnDetail.Buttons.Add(, "recipes", , tbrDefault, "recipes")
    btnX.ToolTipText = "Station Recipe"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    
    'Purge Profiles Screen
    Set btnX = tbrStnDetail.Buttons.Add(, "purgeprofile", , tbrDefault, "purgeprofile")
    btnX.ToolTipText = "Station PurgeProfile"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    
    'Sequence (Courses) Screen
    Set btnX = tbrStnDetail.Buttons.Add(, "courses", , tbrDefault, "courses")
    btnX.ToolTipText = "Station Sequence"
    btnX.Description = btnX.ToolTipText
    
    'Fuel Supply Screen
'    If systemhasLIVEFUEL And systemhasAUTODRAINFILL Then
'        Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
'        Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
        
    If systemhasLIVEFUEL Then
        Set btnX = tbrStnDetail.Buttons.Add(, "fuelsupply", , tbrDefault, "fuelsupply")
        btnX.ToolTipText = "Station Fuel Supply"
        btnX.Description = btnX.ToolTipText
    End If
    
'    Set btnX = tbrStnDetail.Buttons.Add(, , , tbrSeparator)
    
    'Close Screen
'    Set btnX = tbrStnDetail.Buttons.Add(, "close", , tbrDefault, "close")
'    btnX.ToolTipText = "Close This Screen"
'    btnX.Description = btnX.ToolTipText
    
    
    ' Show form to continue configuring
    Show
    
'    With txtDspStn
'        .Height = 0.65 * tbrStnDetail.Buttons("StnNoTxt").Height
'        .Width = 0.5 * tbrStnDetail.Buttons("StnNoTxt").Width
'        .Top = tbrStnDetail.Buttons("StnNoTxt").Top + 15
'        .Left = tbrStnDetail.Buttons("StnNoTxt").Left + (0.25 * tbrStnDetail.Buttons("StnNoTxt").Width)
'        .FontSize = 28
'        .FontBold = False
'        .Locked = True
'    End With
'    With txtStation
'        .Height = 0.2 * tbrStnDetail.Buttons("StnNoTxt").Height
'        .Width = tbrStnDetail.Buttons("StnNoTxt").Width
'        .Top = tbrStnDetail.Buttons("StnNoTxt").Top + (0.7 * tbrStnDetail.Buttons("StnNoTxt").Height)
'        .Left = tbrStnDetail.Buttons("StnNoTxt").Left
'        .FontBold = True
'        .FontSize = 9
'        .text = "Station"
'        .Locked = True
'    End With

'    With txtDspShift
'        .Height = txtDspStn.Height
'        .Width = txtDspStn.Width
'        .Top = txtDspStn.Top
'        .Left = tbrStnDetail.Buttons("ShiftNoTxt").Left + (0.5 * (tbrStnDetail.Buttons("ShiftNoTxt").Width - txtDspShift.Width))
'        .ForeColor = IIf(NR_SHIFT > 1, TitlesData_Forecolor, DKGRAY)
'        .FontSize = 22
'        .FontBold = False
'        .Locked = True
'    End With
'    With txtShift
'        .Height = txtStation.Height
'        .Width = tbrStnDetail.Buttons("ShiftNoTxt").Width
'        .Top = txtStation.Top
'        .Left = tbrStnDetail.Buttons("ShiftNoTxt").Left
'        .ForeColor = IIf(NR_SHIFT > 1, TitlesData_Forecolor, DKGRAY)
'        .FontBold = True
'        .FontSize = 9
'        .text = "Shift"
'        .Locked = True
'    End With

'    tbrStnDetail.Buttons("nextshift").Enabled = IIf(NR_SHIFT > 1, True, False)
'    tbrStnDetail.Buttons("prevshift").Enabled = IIf(NR_SHIFT > 1, True, False)

    tbrStnDetail.Buttons("start").Enabled = IIf(CheckPass("R", False), True, False)
    tbrStnDetail.Buttons("stop").Enabled = IIf(CheckPass("R", False), True, False)
    tbrStnDetail.Buttons("pause").Enabled = IIf(CheckPass("R", False), True, False)

    txtShift.ForeColor = IIf(NR_SHIFT > 1, TitlesData_Forecolor, DKGRAY)
    txtDspShift.ForeColor = IIf(NR_SHIFT > 1, TitlesData_Forecolor, DKGRAY)
    cmdUpShift.Enabled = IIf(NR_SHIFT > 1, True, False)
    cmdDnShift.Enabled = IIf(NR_SHIFT > 1, True, False)
End Sub

Private Sub mnuAbout_Click()
    'About
    menuAbout
End Sub

Private Sub mnuButane_Click()
    menuButane
End Sub

Private Sub mnuCalibration_Click()
    menuCalibration
End Sub

Private Sub mnuCanisters_Click()
    menuCanisters
End Sub

Private Sub mnuConfiguration_Click()
    ' Configuration
    menuConfiguration
End Sub

Private Sub mnuEventLog_Click()
    menuEventLog
End Sub

Private Sub mnuExit_Click()
    ' Exit Program
    menuExit
End Sub

Private Sub mnuIOMonitor_Click()
    menuIoMonitor
End Sub

Private Sub mnuJoblist_Click()
    menuJobList
End Sub

Private Sub mnuLogin_Click()
    menuLogin
End Sub

Private Sub mnuLogout_Click()
    menuLogout
End Sub

Private Sub mnuPurgeProfiles_Click()
    menuPurgeProfiles
End Sub

Private Sub mnuRecipes_Click()
    menuRecipes
End Sub

Private Sub mnuReviewData_Click()
    ' Review Previous Cycle Data
    menuReview
End Sub

Private Sub mnuScaleMonitor_Click()
    menuScaleMonitor
End Sub

Private Sub mnuSysdef_Click()
    ' Select System Definition
    menuSysdef
End Sub

Private Sub mnuWatchData_Click()
    ' Watch Current Cycle Data
    menuWatch
End Sub

Private Sub tbrStnDetail_ButtonClick(ByVal Button As MSComctlLib.Button)
   ' Use the Key property with the SelectCase statement to specify
   ' an action.
   Select Case Button.Key
       Case Is = "prevstn"
            StnDown
       Case Is = "nextstn"
            StnUp
       Case Is = "prevshift"
            ShiftDown
       Case Is = "nextshift"
            ShiftUp
       Case Is = "alarmlog"
            AlarmLog
       Case Is = "ootlog"
            OOTLog
       Case Is = "statsum"
            StatisticsSummary
       Case Is = "joblog"
            JobLog
       Case Is = "opercomment"
            frmOperComment.WhichStn = DispStn
            frmOperComment.WhichShift = DispShift
            frmOperComment.Show
       Case Is = "start"
            If StationControl(DispStn, DispShift).Mode = VBIDLE Then
                If StationCanister(DispStn, DispShift).Validated Then
                    ' start the station
                    StationStart
                Else
                    txtStnDtlMsg.text = vbCrLf & "Must have Valid CANISTER defined first"
                End If
            Else
                StationContinue "Operator"
            End If
       Case Is = "continue"
            StationContinue "Operator"
       Case Is = "pause"
            StationPause "Operator"
       Case Is = "stop"
            StationStop
       Case Is = "canisters"
            frmCanRecipe.Show
            frmCanRecipe.ChgCanRcpMode (CInt(STATIONMODE))
            frmCanRecipe.InitCanRcp
       Case Is = "recipes"
            If StationCanister(DispStn, DispShift).Validated Then
                frmRecipe.Show
                frmRecipe.ChgRecipeMode (CInt(STATIONMODE))
                frmRecipe.InitRecipe
                frmRecipe.Hide
                frmRecipe.Show
            Else
                txtStnDtlMsg.text = vbCrLf & "Must have Valid CANISTER defined first"
            End If
       Case Is = "purgeprofile"
            If (StationRecipe(DispStn, DispShift).Purge_Method = PURGEBYPROFILE) Then
                frmPurgeProfile.Show
                frmPurgeProfile.ChgProfileMode (CInt(STATIONMODE))
                frmPurgeProfile.InitProfile
            Else
                txtStnDtlMsg.text = "Current Recipe does not use Purge-By-Profile"
            End If
       Case Is = "courses"
            JobSeqAutoEdit = True
            frmCourses.Show
            frmCourses.ChgJobSeqMode (CInt(STATIONMODE))
            frmCourses.InitSeqRcp
       Case Is = "fuelsupply"
            'Fuel Supply Screen
            If ((STN_INFO(DispStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE)) Then
                If STN_INFO(DispStn).ADF_TANKTYPE <> 0 Then
                    frmFuelSupply.Show
                Else
                    txtStnDtlMsg.text = "Current Station does not have ADF Control"
                End If
            Else
                txtStnDtlMsg.text = "Current Station does not support Live Fuel"
            End If
    
    '   Case Is = "close"
    '        CloseScreen
   End Select
End Sub

Private Sub tbrNavigate_ButtonClick(ByVal Button As MSComctlLib.Button)
   ' Use the Key property with the SelectCase statement to specify
   ' an action.
   Select Case Button.Key
       Case Is = "overview"
            menuOverview
       Case Is = "stndetail"
            menuStnDetail
       Case Is = "reviewdata"
            menuReview
       Case Is = "watchdata"
            ' Watch Current Cycle Data
            menuWatch
       Case Is = "login"
            menuLogin
       Case Is = "logout"
            menuLogout
       Case Is = "copyfiles"
            ' CopyFiles
            menuCopyFile
       Case Is = "printfiles"
            ' PrintFiles
            menuPrintFile
       Case Is = "butane"
            menuButane
       Case Is = "fueluselog"
            menuViewFuelUseLog
       Case Is = "canisters"
            menuCanisters
       Case Is = "recipes"
            menuRecipes
       Case Is = "courses"
            menuCourses
       Case Is = "purgeprofile"
            menuPurgeProfiles
       Case Is = "tomcanload"
            menuRemCanLoad
       Case Is = "configuration"
            ' Configuration
            menuConfiguration
       Case Is = "sysdef"
            ' Select System Definition
            menuSysdef
       Case Is = "eventlog"
            menuEventLog
       Case Is = "joblist"
            menuJobList
       Case Is = "calibration"
            menuCalibration
       Case Is = "iomonitor"
            menuIoMonitor
       Case Is = "scalemonitor"
            menuScaleMonitor
       Case Is = "leaktest"
            menuLeakTest
       Case Is = "simulation"
            ' Simulation
            menuSimulation
       Case Is = "opermanual"
            ' Operators Manual
            menuOperatorManual
       Case Is = "beaconoff"
            ' TurnOff Beacon
            menuBeaconOff
       Case Is = "hornoff"
            ' TurnOff Horn
            menuHornOff
       Case Is = "firstaid"
            ' First Aid File Save
            menuFirstAid
'       Case Is = "close"
            ' Close Screen
'            CloseScreen
   End Select
End Sub

Private Sub UpdateStatusBars()
    ' Status Bar #1
    pnlEstop.BackColor = frmMainMenu.pnlEstop.BackColor
    pnlFlow.BackColor = frmMainMenu.pnlFlow.BackColor
    pnlBtn20.BackColor = frmMainMenu.pnlBtn20.BackColor
    pnlDoors.BackColor = frmMainMenu.pnlDoors.BackColor
    pnlComms.BackColor = frmMainMenu.pnlComms.BackColor
    pnlEstop.ToolTipText = frmMainMenu.pnlEstop.ToolTipText
    pnlFlow.ToolTipText = frmMainMenu.pnlFlow.ToolTipText
    pnlBtn20.ToolTipText = frmMainMenu.pnlBtn20.ToolTipText
    pnlDoors.ToolTipText = frmMainMenu.pnlDoors.ToolTipText
    pnlComms.ToolTipText = frmMainMenu.pnlComms.ToolTipText
    pnlMix.BackColor = frmMainMenu.pnlMix.BackColor
    pnlMix.ToolTipText = frmMainMenu.pnlMix.ToolTipText
    pnlMix.Top = frmMainMenu.pnlMix.Top
    pnlMessage.Font = frmMainMenu.pnlMessage.Font
    pnlMessage.FontSize = frmMainMenu.pnlMessage.FontSize
    pnlMessage.BackColor = SysMessage_BackColor
    pnlMessage.ForeColor = SysMessage_ForeColor
    pnlMessage.Caption = SysMessage_Text
    pnlMessage.ToolTipText = SysMessage_Tooltip
    pnlPurgeAir.ForeColor = PurgeAirMsg_ForeColor
    pnlPurgeAir.Caption = PurgeAirMsg_Text
    pnlPurgeAir.ToolTipText = PurgeAirMsg_ToolTip
End Sub

Sub UpdateNavigateBtns()

'
' Routine Name:  UpdateNavigateBtns
' Description:
' Updates the Navigate toolbar buttons
'
Dim iKeyCount As Integer
 
SetErrModule 89, 10101
If UseLocalErrorHandler Then On Error GoTo localhandler
        
        ' Login
        If CheckPass("J", False) Then
'            tbrNavigate.Buttons("login").Enabled = True
            mnuLogin.Enabled = True
        Else
'            tbrNavigate.Buttons("login").Enabled = False
            mnuLogin.Enabled = False
        End If
        
        ' Logout
'        tbrNavigate.Buttons("logout").Enabled = True
        mnuLogout.Enabled = True
        
        ' CopyFiles
        If CheckPass("F", False) Then
'            tbrNavigate.Buttons("copyfiles").Visible = True
'            tbrNavigate.Buttons("copyfiles").Enabled = True
            mnuCopyFile.Enabled = True
        Else
'            tbrNavigate.Buttons("copyfiles").Visible = False
'            tbrNavigate.Buttons("copyfiles").Enabled = False
            mnuCopyFile.Enabled = False
        End If

        ' PrintFiles
        If CheckPass("F", False) Then
'            tbrNavigate.Buttons("printfiles").Visible = True
'            tbrNavigate.Buttons("printfiles").Enabled = True
            mnuPrintFile.Enabled = True
        Else
'            tbrNavigate.Buttons("printfiles").Visible = False
'            tbrNavigate.Buttons("printfiles").Enabled = False
            mnuPrintFile.Enabled = False
        End If
        
        ' Canisters
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("canisters").Enabled = True
            mnuCanisters.Enabled = True
        Else
            tbrNavigate.Buttons("canisters").Enabled = False
            mnuCanisters.Enabled = False
        End If
        
        ' Recipes
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("recipes").Enabled = True
            mnuRecipes.Enabled = True
        Else
            tbrNavigate.Buttons("recipes").Enabled = False
            mnuRecipes.Enabled = False
        End If
        
        ' Purge Profiles
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("purgeprofile").Enabled = True
            mnuPurgeProfiles.Enabled = True
        Else
            tbrNavigate.Buttons("purgeprofile").Enabled = False
            mnuPurgeProfiles.Enabled = False
        End If
        
        ' Courses
        If CheckPass("N", False) And (NR_JOBSEQ > 1) Then
            tbrNavigate.Buttons("courses").Visible = True
            mnuCourses.Visible = True
        Else
            tbrNavigate.Buttons("courses").Visible = False
            mnuCourses.Visible = False
        End If
        
        ' TomCanLoad
        If CheckPass("N", False) And (USINGREMCANLOAD Or USINGTOMCANLOAD) Then
            tbrNavigate.Buttons("tomcanload").Visible = True
            mnuTomCanLoad.Visible = True
        Else
            tbrNavigate.Buttons("tomcanload").Visible = False
            mnuTomCanLoad.Visible = False
        End If
        
        ' Configuration
        If CheckPass("B", False) Then
            tbrNavigate.Buttons("configuration").Enabled = True
            mnuConfiguration.Enabled = True
        Else
            mnuConfiguration.Enabled = False
            tbrNavigate.Buttons("configuration").Enabled = False
        End If
        
        ' System Definition
        If CheckPass("H", False) Then
            tbrNavigate.Buttons("sysdef").Visible = True
            tbrNavigate.Buttons("sysdef").ToolTipText = "System Definition"
            mnuSysdef.Visible = True
        Else
            tbrNavigate.Buttons("sysdef").Visible = False
            tbrNavigate.Buttons("sysdef").ToolTipText = ""
            mnuSysdef.Visible = False
        End If
        
        ' Butane Available
        If systemhasBUTANE Then
            tbrNavigate.Buttons("butane").Visible = True
            mnuButane.Enabled = True
        Else
            tbrNavigate.Buttons("butane").Visible = False
            mnuButane.Enabled = False
        End If
        
        ' Event Log
        If CheckPass("Z", False) Then
            tbrNavigate.Buttons("eventlog").Enabled = True
            mnuEventLog.Enabled = True
        Else
            mnuEventLog.Enabled = False
            tbrNavigate.Buttons("eventlog").Enabled = False
        End If
        
        ' Joblist Log
        If CheckPass("M", False) Then
            tbrNavigate.Buttons("joblist").Enabled = True
            mnuJoblist.Enabled = True
        Else
            mnuJoblist.Enabled = False
            tbrNavigate.Buttons("joblist").Enabled = False
        End If
        
        ' Review Previous Cycle Data
        If CheckPass("F", False) Then
            tbrNavigate.Buttons("reviewdata").Enabled = True
            mnuReviewData.Enabled = True
        Else
            mnuReviewData.Enabled = False
            tbrNavigate.Buttons("reviewdata").Enabled = False
        End If
        
        ' Watch Current Cycle Data
        If CheckPass("F", False) Then
            tbrNavigate.Buttons("watchdata").Enabled = True
            mnuWatchData.Enabled = True
        Else
            mnuWatchData.Enabled = False
            tbrNavigate.Buttons("watchdata").Enabled = False
        End If
        
        ' MFC Calibration
        If CheckPass("X", False) Then
            tbrNavigate.Buttons("calibration").Enabled = True
            mnuCalibration.Enabled = True
        Else
            mnuCalibration.Enabled = False
            tbrNavigate.Buttons("calibration").Enabled = False
        End If
        
        ' I/O Monitor
        If CheckPass("2", False) Then
            tbrNavigate.Buttons("iomonitor").Enabled = True
            mnuIomonitor.Enabled = True
        Else
            mnuIomonitor.Enabled = False
            tbrNavigate.Buttons("iomonitor").Enabled = False
        End If
        
        ' Scale Monitor
        If CheckPass("3", False) Then
            tbrNavigate.Buttons("scalemonitor").Enabled = True
            mnuScaleMonitor.Enabled = True
        Else
            mnuScaleMonitor.Enabled = False
            tbrNavigate.Buttons("scalemonitor").Enabled = False
        End If
              
        ' Simulation
        If Not IoComOn And USINGSIMULATION And CheckPass("H", False) Then
            tbrNavigate.Buttons("simulation").Visible = True
            tbrNavigate.Buttons("simulation").Enabled = True
            tbrNavigate.Buttons("simulation").ToolTipText = "Simulation Control Panel"
'            mnuSimulation.Enabled = True
        Else
'            mnuSimulation.Enabled = False
            tbrNavigate.Buttons("simulation").Visible = False
            tbrNavigate.Buttons("simulation").Enabled = False
            tbrNavigate.Buttons("simulation").ToolTipText = ""
        End If
        
        ' Operator Manual
        If CheckPass("H", False) Then
            tbrNavigate.Buttons("opermanual").Visible = False
            tbrNavigate.Buttons("opermanual").Enabled = False
            mnuOperatorManual.Enabled = True
        ElseIf CheckPass("D", False) Then
            tbrNavigate.Buttons("opermanual").Visible = True
            tbrNavigate.Buttons("opermanual").Enabled = True
            mnuOperatorManual.Enabled = True
        Else
            mnuOperatorManual.Enabled = False
            tbrNavigate.Buttons("opermanual").Visible = False
            tbrNavigate.Buttons("opermanual").Enabled = False
        End If
        
'        ' FirstAid
        If CheckPass("T", False) Then
'            tbrNavigate.Buttons("firstaid").Visible = True
'            tbrNavigate.Buttons("firstaid").Enabled = True
'            tbrNavigate.Buttons("firstaid").ToolTipText = "FirstAid File Save for APS"
            mnuFirstAid.Enabled = True
        Else
            mnuFirstAid.Enabled = False
'            tbrNavigate.Buttons("firstaid").Visible = False
'            tbrNavigate.Buttons("firstaid").Enabled = False
'            tbrNavigate.Buttons("firstaid").ToolTipText = ""
        End If
        
        ' Close Screen
'        tbrNavigate.Buttons("close").Enabled = True
        
        ' View AirLog
        If LogTempRh Then
            mnuAirLog.Enabled = True
        Else
            mnuAirLog.Enabled = False
        End If
        
        ' Exit Program
        If CheckPass("G", False) Then
'            tbrNavigate.Buttons("exit").Enabled = True
            mnuExit.Enabled = True
        Else
            mnuExit.Enabled = False
'            tbrNavigate.Buttons("exit").Enabled = False
        End If

        
        ' **********************
        ' **********************
        ' STATION DETAIL TOOLBAR
        ' **********************
        ' **********************
        
        tbrStnDetail.Buttons("alarmlog").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrStnDetail.Buttons("ootlog").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrStnDetail.Buttons("statsum").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrStnDetail.Buttons("joblog").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrStnDetail.Buttons("opercomment").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrStnDetail.Buttons("opercomment").Visible = True
        
        ' Station PurgeProfile
        If (StationRecipe(DispStn, DispShift).Purge_Method = PURGEBYPROFILE) Then
            tbrStnDetail.Buttons("purgeprofile").Enabled = True
        Else
            tbrStnDetail.Buttons("purgeprofile").Enabled = False
        End If
        
        ' Station Courses
        If (NR_JOBSEQ > 1) Then
            tbrStnDetail.Buttons("courses").Visible = True
        Else
            tbrStnDetail.Buttons("courses").Visible = False
        End If
        
        'Fuel Supply Screen
        If systemhasLIVEFUEL Then
            If ((STN_INFO(DispStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE)) Then
                If STN_INFO(DispStn).ADF_TANKTYPE <> 0 Then
                    tbrStnDetail.Buttons("fuelsupply").Enabled = True
                Else
                    tbrStnDetail.Buttons("fuelsupply").Enabled = False
                End If
            Else
                tbrStnDetail.Buttons("fuelsupply").Enabled = False
            End If
        End If
    
        
        '*************************************
        '*************************************
        '*************************************
        ' START/CONTINUE, PAUSE & STOP BUTTONS
        ' START/CONTINUE, PAUSE & STOP BUTTONS
        ' START/CONTINUE, PAUSE & STOP BUTTONS
        '*************************************
        '*************************************
        '*************************************
        ChgErrModule 89, 10102
        If Stop_In_Progress = True Then
            tbrStnDetail.Buttons("start").Enabled = False
            If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
            If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
            tbrStnDetail.Buttons("stop").Enabled = False
        Else
            Select Case StationControl(DispStn, DispShift).Mode
                Case VBIDLE
                    tbrStnDetail.Buttons("continue").Visible = False
                    tbrStnDetail.Buttons("pause").Visible = False
                    tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = False
                    tbrStnDetail.Buttons("start").Visible = True
                    tbrStnDetail.Buttons("start").ToolTipText = "Start Job"
                    If AdfControl(DispStn).Step = 0 Then
                        If Pause_Alarm = SYSTEMPAUSED Then
                            ' system is paused
                            tbrStnDetail.Buttons("start").Enabled = False
                        Else
                            ' station is not paused
                            If Stn_OperReportNameIsValid = False Then
                                If (SysConfig.ReportFileName1stPart = RPT_OPERENTER _
                                  Or SysConfig.ReportFileName2ndPart = RPT_OPERENTER _
                                  Or SysConfig.ReportFileName3rdPart = RPT_OPERENTER) Then
                                ' operator entry of report name required before Job can be started
                                tbrStnDetail.Buttons("start").Enabled = False
                                Else
                                ' no operator entry required
                                tbrStnDetail.Buttons("start").Enabled = True
                                End If
                            Else
                                ' valid operator entry
                                tbrStnDetail.Buttons("start").Enabled = True
                            End If
                        End If
                    Else
                        tbrStnDetail.Buttons("start").Enabled = False
                    End If
            
                Case VBLEAKERROR
                    tbrStnDetail.Buttons("continue").Visible = False
                    tbrStnDetail.Buttons("pause").Visible = False
                    tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    tbrStnDetail.Buttons("start").Visible = True
                    Select Case SysConfig.LeakCheckFailResponse
                        Case MANUALCHOOSE, AUTOCONTINUE
                            tbrStnDetail.Buttons("start").Enabled = True
                            tbrStnDetail.Buttons("start").ToolTipText = "CONTINUE"
                        Case Else
                            tbrStnDetail.Buttons("start").Enabled = False
                            tbrStnDetail.Buttons("start").ToolTipText = "CONTINUE"
                    End Select
            
                Case VBPAUSEALARM
                    tbrStnDetail.Buttons("continue").Visible = False
                    tbrStnDetail.Buttons("pause").Visible = False
                    If Pause_Alarm = SYSTEMPAUSED Then
                        ' System is Paused
                        tbrStnDetail.Buttons("start").Visible = True
                        tbrStnDetail.Buttons("start").Enabled = False
                        tbrStnDetail.Buttons("stop").Visible = True
                        tbrStnDetail.Buttons("stop").Enabled = False
                    Else
                        ' Station is Paused
                        tbrStnDetail.Buttons("start").Visible = True
                        tbrStnDetail.Buttons("start").Enabled = True
                        tbrStnDetail.Buttons("stop").Visible = True
                        tbrStnDetail.Buttons("stop").Enabled = True
                        tempText = "Continue"
                        If StationControl(DispStn, DispShift).Mode_PauseSave = VBLEAK Then tempText = "Restart Leak Check"
                        If StationControl(DispStn, DispShift).Mode_PauseSave = VBLOAD Then tempText = "Resume Load"
                        If StationControl(DispStn, DispShift).Mode_PauseSave = VBPURGE Then tempText = "Resume Purge"
                        tbrStnDetail.Buttons("start").ToolTipText = tempText
                    End If
            
                Case VBPAUSEOOT
                    tbrStnDetail.Buttons("continue").Visible = False
                    tbrStnDetail.Buttons("pause").Visible = False
                    tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    tbrStnDetail.Buttons("start").Visible = True
                    Select Case StationControl(DispStn, DispShift).Mode_PauseSave
                        Case VBLEAK
                            tbrStnDetail.Buttons("start").Enabled = False
                            tbrStnDetail.Buttons("start").ToolTipText = "Continue"
                        Case VBLOAD
                            tbrStnDetail.Buttons("start").Enabled = IIf((StationControl(DispStn, DispShift).OotResponse = ootrspStop), False, True)
                            tbrStnDetail.Buttons("start").ToolTipText = "Resume Load"
                        Case VBPURGE, VBPURGECONT
                            tbrStnDetail.Buttons("start").Enabled = IIf((StationControl(DispStn, DispShift).OotResponse = ootrspStop), False, True)
                            tbrStnDetail.Buttons("start").ToolTipText = "Resume Purge"
                        Case Else
                            tbrStnDetail.Buttons("start").Enabled = IIf((StationControl(DispStn, DispShift).OotResponse = ootrspStop), False, True)
                            tbrStnDetail.Buttons("start").ToolTipText = "Continue"
                    End Select
                    
                Case VBFIDPAUSE                                  ' Pause for FID
                    If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If Not tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = True
                    tbrStnDetail.Buttons("start").Enabled = True
                    tbrStnDetail.Buttons("start").ToolTipText = "Continue; FID is ready"

                Case VBGASPAUSE                                  ' Pause for Live Fuel
                    If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If Not tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = True
'                    If (STN_INFO(DispStn).ADF_DEF.hasADF_WaterBath And StationRecipe(DispStn, DispShift).ADF_Heater) Then
'                        If (LoadControl(DispStn, DispShift).WaterBathTempOK) Then
'                            tbrStnDetail.Buttons("start").Enabled = True
'                            tbrStnDetail.Buttons("start").ToolTipText = "Continue; Vapor Tank is ready"
'                        Else
'                            tbrStnDetail.Buttons("start").Enabled = False
'                            tbrStnDetail.Buttons("start").ToolTipText = "Vapor Tank is Not Ready; check WaterBath"
'                        End If
'                    Else
                        tbrStnDetail.Buttons("start").Enabled = True
                        tbrStnDetail.Buttons("start").ToolTipText = "Continue; Vapor Tank is ready"
'                    End If
          
                Case VBWBPAUSE                                  ' Pause for WatterBath
                    If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If Not tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = True
                    tbrStnDetail.Buttons("start").Enabled = True
                    tbrStnDetail.Buttons("start").ToolTipText = "Continue; WaterBath is ready"
          
                Case VBPOSTLOADOPER                              ' PostLoad Pause for Operator
                    If (Not tbrStnDetail.Buttons("continue").Visible) Then tbrStnDetail.Buttons("continue").Visible = True
                    tbrStnDetail.Buttons("continue").Enabled = True
                    tbrStnDetail.Buttons("continue").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = False
          
                Case VBPOSTPURGEOPER                              ' PostPurge Pause for Operator
                    If (Not tbrStnDetail.Buttons("continue").Visible) Then tbrStnDetail.Buttons("continue").Visible = True
                    tbrStnDetail.Buttons("continue").Enabled = True
                    tbrStnDetail.Buttons("continue").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = False
          
                Case VBPURGEWAIT
                    tbrStnDetail.Buttons("continue").Visible = False
                    tbrStnDetail.Buttons("pause").Visible = False
                    tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    tbrStnDetail.Buttons("start").Visible = True
                    If USINGPASLOCALCONTROL And PAS_INFO(pasTEMPERATURE).timeOut Then
                        ' Local PAS Temperature Control Timeout
                        tbrStnDetail.Buttons("start").Enabled = True
                        tbrStnDetail.Buttons("start").ToolTipText = "Reset PAS Temperature Timeout"
                    ElseIf USINGPASLOCALCONTROL And PAS_INFO(pasMOISTURE).timeOut Then
                        ' Local PAS Moisture Control Timeout
                        tbrStnDetail.Buttons("start").Enabled = True
                        tbrStnDetail.Buttons("start").ToolTipText = "Reset PAS Moisture Timeout"
                    Else
                        ' Not using local PAS control or no timeouts
                        tbrStnDetail.Buttons("start").Enabled = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).Enabled
                        tbrStnDetail.Buttons("start").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
                    End If
            
                Case VBPAUSEVACSW                                   ' System Vacuum Switch Off; Wait for Resume from Operator after Vacuum Switch is On
                    If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If Not tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = True
                    tbrStnDetail.Buttons("start").Enabled = IIf(Alm_SystemVacSw, False, True)
                    tbrStnDetail.Buttons("start").ToolTipText = IIf(Alm_SystemVacSw, "Cannot Resume until System Vacuum Switch is True", "Resume; System Vacumm Switch is now True")
                
                Case VBPAUSEBYUSER
                    If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If Not tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = True
                    tbrStnDetail.Buttons("start").Enabled = True
                    tbrStnDetail.Buttons("start").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
            
                Case VBCOURSEWAIT
                    If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
                    If tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = False
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If Not tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = True
                    tbrStnDetail.Buttons("start").Enabled = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).Enabled
                    tbrStnDetail.Buttons("start").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
            
                Case Else
                    If tbrStnDetail.Buttons("continue").Visible Then tbrStnDetail.Buttons("continue").Visible = False
                    If Not tbrStnDetail.Buttons("pause").Visible Then tbrStnDetail.Buttons("pause").Visible = True
                    tbrStnDetail.Buttons("pause").Enabled = True
                    If Not tbrStnDetail.Buttons("stop").Visible Then tbrStnDetail.Buttons("stop").Visible = True
                    tbrStnDetail.Buttons("stop").Enabled = True
                    If tbrStnDetail.Buttons("start").Visible Then tbrStnDetail.Buttons("start").Visible = False
                    tbrStnDetail.Buttons("start").Enabled = False
                    tbrStnDetail.Buttons("start").ToolTipText = ""
            
            End Select
            ' no toolTip if button is not enabled
            tbrStnDetail.Buttons("start").ToolTipText = IIf(tbrStnDetail.Buttons("start").Enabled, tbrStnDetail.Buttons("start").ToolTipText, "")
            
        End If
        
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Public Sub ChartXYValues(ByVal stn As Integer, ByVal Shift As Integer)
Dim i As Integer
    For i = 1 To NumPoints
        Graph(i, 1) = StnGraph(stn, Shift, i, 1)                     ' value for X-axis
        Graph(i, 2) = StnGraph(stn, Shift, i, 2)
        Graph(i, 3) = StnGraph(stn, Shift, i, 3)                     ' value for X-axis
        Graph(i, 4) = StnGraph(stn, Shift, i, 4)
        Graph(i, 5) = StnGraph(stn, Shift, i, 5)                     ' value for X-axis
        Graph(i, 6) = StnGraph(stn, Shift, i, 6)
    Next i
    chtStnChart = Graph ' populate chart's data grid using Graph array
    chtStnChart.Column = 1
    chtStnChart.ColumnLabel = "MFC Mass Flow"
    chtStnChart.Column = 3
    chtStnChart.ColumnLabel = "Primary Scale"
    chtStnChart.Column = 5
    chtStnChart.ColumnLabel = "Aux. Scale"
    chtStnChart.Repaint = True
End Sub

Public Sub DataCollector(ByVal stn As Integer, ByVal Shift As Integer)
Dim i As Integer
    For i = 1 To (NumPoints - 1)
        StnGraph(stn, Shift, i, 1) = i                     ' value for X-axis
        StnGraph(stn, Shift, i, 2) = StnGraph(stn, Shift, i + 1, 2)
        StnGraph(stn, Shift, i, 3) = i                     ' value for X-axis
        StnGraph(stn, Shift, i, 4) = StnGraph(stn, Shift, i + 1, 4)
        StnGraph(stn, Shift, i, 5) = i                     ' value for X-axis
        StnGraph(stn, Shift, i, 6) = StnGraph(stn, Shift, i + 1, 6)
    Next i
    StnGraph(stn, Shift, NumPoints, 1) = NumPoints                     ' value for X-axis
    StnGraph(stn, Shift, NumPoints, 2) = LoadControl(stn, Shift).loadTotalGrams
    StnGraph(stn, Shift, NumPoints, 3) = NumPoints                     ' value for X-axis
    StnGraph(stn, Shift, NumPoints, 4) = StationControl(stn, Shift).PriScaleWt - Stn_PriScale_RefValues(stn, Shift)
    StnGraph(stn, Shift, NumPoints, 5) = NumPoints                     ' value for X-axis
    StnGraph(stn, Shift, NumPoints, 6) = StationControl(stn, Shift).AuxScaleWt - Stn_AuxScale_RefValues(stn, Shift)
    NumPointsSoFar(stn, Shift) = NumPointsSoFar(stn, Shift) + 1
    ' check for out of scale range
    ' any current value above current scale max ??
    If (StnGraph(stn, Shift, NumPoints, 2) > chtStnChart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum) Then
        Scale_Yaxis StnGraph(stn, Shift, NumPoints, 2)
    ElseIf (StnGraph(stn, Shift, NumPoints, 4) > chtStnChart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum) Then
        Scale_Yaxis StnGraph(stn, Shift, NumPoints, 4)
    ElseIf (StnGraph(stn, Shift, NumPoints, 6) > chtStnChart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum) Then
        Scale_Yaxis StnGraph(stn, Shift, NumPoints, 6)
    End If
End Sub

Public Sub ClearXYvalues(ByVal stn As Integer, ByVal Shift As Integer)
Dim i As Integer
    For i = 1 To NumPoints
        StnGraph(stn, Shift, i, 1) = 0
        StnGraph(stn, Shift, i, 2) = 0
        StnGraph(stn, Shift, i, 3) = 0
        StnGraph(stn, Shift, i, 4) = 0
        StnGraph(stn, Shift, i, 5) = 0
        StnGraph(stn, Shift, i, 6) = 0
    Next i
    NumPointsSoFar(stn, Shift) = 0
    If (stn = DispStn) And (Shift = DispShift) Then ChartXYValues stn, Shift
End Sub

Public Sub Shift_Yvalues(ByVal stn As Integer, ByVal Shift As Integer)
Dim i As Integer
Dim Idx As Integer
Dim idx2 As Integer
    Idx = IIf((NumPointsSoFar(stn, Shift) < NumPoints), (1 + NumPoints - NumPointsSoFar(stn, Shift)), 1)
    idx2 = IIf((NumPointsSoFar(stn, Shift) < NumPoints), (Idx + 150), 1)
    For i = Idx To NumPoints
        If i >= idx2 Then
            StnGraph(stn, Shift, i, 4) = StnGraph(stn, Shift, i, 4) - StnGraph(stn, Shift, NumPoints, 4)
            StnGraph(stn, Shift, i, 6) = StnGraph(stn, Shift, i, 6) - StnGraph(stn, Shift, NumPoints, 6)
        Else
            StnGraph(stn, Shift, i, 2) = 0
            StnGraph(stn, Shift, i, 4) = 0
            StnGraph(stn, Shift, i, 6) = 0
        End If
    Next i
    If (stn = DispStn) And (Shift = DispShift) Then ChartXYValues stn, Shift
End Sub

Public Sub Scale_Yaxis(ByVal CanWC As Single)
Dim NewSpan, TargetSpan, mult As Single
Dim NewDiv, newmax, NewMin As Single
Dim NewMajDiv, NewMinDiv As Integer
Dim NewAuto As Boolean
Dim clr As Long
Dim canVol As Single

    canVol = StationCanister(DispStn, DispShift).WorkingVolume
    TargetSpan = IIf((CanWC > 0), (1.28 * CanWC), (15.7 * canVol))
    
    mult = 1
    Do While TargetSpan > 10
        TargetSpan = TargetSpan / 10
        mult = mult * 10
    Loop
            
    If mult > 1000 Then
        NewAuto = True
    ElseIf mult > 1 Then
        Select Case TargetSpan
            Case Is > 7
                NewSpan = 12 * mult
                newmax = 10 * mult
                NewMin = -2 * mult
                NewDiv = 1 * mult
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 5
                NewSpan = 9 * mult
                newmax = 7.5 * mult
                NewMin = -1.5 * mult
                NewDiv = 0.75 * mult
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 4
                NewSpan = 6 * mult
                newmax = 5 * mult
                NewMin = -1 * mult
                NewDiv = 0.5 * mult
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 3
                NewSpan = 4.8 * mult
                newmax = 4 * mult
                NewMin = -0.8 * mult
                NewDiv = 0.4
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 2
                NewSpan = 3.6 * mult
                newmax = 3 * mult
                NewMin = -0.6 * mult
                NewDiv = 0.3
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 1
                NewSpan = 2.4 * mult
                newmax = 2 * mult
                NewMin = -0.4 * mult
                NewDiv = 0.2
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Else
                NewSpan = 1.2 * mult
                newmax = 1 * mult
                NewMin = -0.2 * mult
                NewDiv = 0.1
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
        End Select
    Else
        Select Case TargetSpan
            Case Is > 7
                NewSpan = 12 * mult
                newmax = 10 * mult
                NewMin = -2 * mult
                NewDiv = 1 * mult
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 5
                NewSpan = 9 * mult
                newmax = 7.5 * mult
                NewMin = -1.5 * mult
                NewDiv = 0.75 * mult
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 4
                NewSpan = 6 * mult
                newmax = 5 * mult
                NewMin = -1 * mult
                NewDiv = 0.5 * mult
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 3
                NewSpan = 4.5 * mult
                newmax = 4 * mult
                NewMin = -0.5 * mult
                NewDiv = 0.4
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 2
                NewSpan = 3.5 * mult
                newmax = 3 * mult
                NewMin = -0.5 * mult
                NewDiv = 0.3
                NewMajDiv = 6
                NewMinDiv = 2
                NewAuto = False
            Case Is > 1
                NewSpan = 2.5 * mult
                newmax = 2 * mult
                NewMin = -0.5 * mult
                NewDiv = 0.2
                NewMajDiv = 5
                NewMinDiv = 2
                NewAuto = False
            Case Else
                NewSpan = 2# * mult
                newmax = 1.5 * mult
                NewMin = -0.5 * mult
                NewDiv = 0.15
                NewMajDiv = 4
                NewMinDiv = 2
                NewAuto = False
        End Select
    End If
    

    With chtStnChart.Plot
    
        ' Y axis
        With .Axis(VtChAxisIdY).ValueScale
            .Auto = NewAuto
            If Not NewAuto Then .Minimum = NewMin
            If Not NewAuto Then .Maximum = newmax
            If Not NewAuto Then .MajorDivision = NewMajDiv
            If Not NewAuto Then .MinorDivision = NewMinDiv
        End With
        
        ' X axis
        With .Axis(VtChAxisIdX).ValueScale
            .Auto = False
            .Minimum = 0
            .Maximum = 1000
            .MajorDivision = 10
            .MinorDivision = 2
        End With
        
        ' LoadTotal Pen
        clr = DK2GREEN
        .SeriesCollection(1).Pen.Style = VtPenStyleSolid
        .SeriesCollection(1).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
        
        ' Primary Scale Pen
        clr = MEDBLUE
        .SeriesCollection(3).Pen.Style = VtPenStyleSolid
        .SeriesCollection(3).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
        
        ' Aux Scale Pen
        clr = MEDORANGE
        .SeriesCollection(5).Pen.Style = VtPenStyleSolid
        .SeriesCollection(5).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
        
    End With
    
End Sub

Public Sub SetXYtimeInterval(ByVal stn As Integer, ByVal Shift As Integer)
Dim dVal, dVal1 As Double
    ' x-axis time interval
    dVal1 = CDbl(60) * StationControl(stn, Shift).EstJobDur
    dVal = dVal1 / NumPoints
    If dVal > CDbl(600) Then dVal = CDbl(600)
    If dVal < CDbl(2) Then dVal = CDbl(2)
    Stn_XYChart_Xinterval(stn, Shift) = dVal
End Sub

Public Sub SlideXYgraph(ByVal stn As Integer, ByVal Shift As Integer)
Dim iPoints As Integer
    ' Slide stn XY Graph to the left by 1 time division
    For iPoints = 1 To 50
        DataCollector stn, Shift
    Next iPoints
    If (stn = DispStn) And (Shift = DispShift) Then ChartXYValues stn, Shift
End Sub
    
Public Sub RemoteStnStart(ByVal stn As Integer, ByVal Shift As Integer)
    DispStn = stn
    DispShift = Shift
    StationStart
End Sub
