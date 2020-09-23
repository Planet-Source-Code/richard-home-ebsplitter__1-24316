VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EBSplitterTest"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdMulti 
      Caption         =   "MultiSplitter Demo"
      Height          =   555
      Left            =   420
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "Text File Browser"
      Height          =   555
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowser_Click()

    frmExplorer.Show vbModal, Me
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdMulti_Click()

    frmMulti.Show vbModal, Me
    
    
End Sub

Private Sub lblPrompt_Click()

End Sub
