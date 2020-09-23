VERSION 5.00
Object = "*\AEBSplitter.vbp"
Begin VB.Form frmMulti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EBSplitter Multi Splitter Demo"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin EBSplitter.ctlEBSplitter ctlEBSplitter 
      Height          =   3015
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   5318
      SplitterPos     =   968
      Begin EBSplitter.ctlEBSplitter ctlEBSplitter1 
         Height          =   3015
         Left            =   1028
         TabIndex        =   2
         Top             =   0
         Width           =   3367
         _ExtentX        =   5927
         _ExtentY        =   5318
         Orientation     =   0
         SplitterPos     =   1478
         Begin EBSplitter.ctlEBSplitter ctlEBSplitter2 
            Height          =   1477
            Left            =   0
            TabIndex        =   4
            Top             =   1538
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   2593
            SplitterPos     =   1162
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00C00000&
               Height          =   1470
               Left            =   1222
               ScaleHeight     =   1410
               ScaleWidth      =   2085
               TabIndex        =   6
               Top             =   0
               Width           =   2138
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H0000C000&
               Height          =   1470
               Left            =   0
               ScaleHeight     =   1410
               ScaleWidth      =   1110
               TabIndex        =   5
               Top             =   0
               Width           =   1177
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H0000C0C0&
            Height          =   1493
            Left            =   0
            ScaleHeight     =   1440
            ScaleWidth      =   3300
            TabIndex        =   3
            Top             =   0
            Width           =   3360
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000C0&
         Height          =   3015
         Left            =   0
         ScaleHeight     =   2955
         ScaleWidth      =   930
         TabIndex        =   1
         Top             =   0
         Width           =   983
      End
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

