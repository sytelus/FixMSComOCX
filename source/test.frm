VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   810
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1429
      _CBWidth        =   3615
      _CBHeight       =   810
      _Version        =   "6.0.8141"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1440
      NewRow2         =   -1  'True
      MinHeight3      =   360
      Width3          =   1440
      NewRow3         =   0   'False
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      _Version        =   393216
      Enabled         =   -1  'True
      FullWidth       =   73
      FullHeight      =   41
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   1296
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

