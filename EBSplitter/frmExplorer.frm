VERSION 5.00
Object = "*\AEBSplitter.vbp"
Begin VB.Form frmExplorer 
   Caption         =   "Text File Browser"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin EBSplitter.ctlEBSplitter ctlSplitterMain 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4683
      SplitterPos     =   968
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   1028
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   0
         Width           =   3307
      End
      Begin VB.PictureBox picExplorer 
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2595
         ScaleWidth      =   930
         TabIndex        =   1
         Top             =   0
         Width           =   983
         Begin VB.DriveListBox drv 
            Height          =   315
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   915
         End
         Begin EBSplitter.ctlEBSplitter ctlSplitterFolderFile 
            Height          =   2175
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   3836
            Orientation     =   0
            SplitterPos     =   1088
            Begin VB.FileListBox fil 
               Height          =   870
               Left            =   0
               Pattern         =   "*.txt"
               TabIndex        =   6
               Top             =   1148
               Width           =   915
            End
            Begin VB.DirListBox dir 
               Height          =   990
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   915
            End
         End
      End
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[Description]
'   Quick example of the EBSplitter control
'   This form uses two splitters (one contained in the other) to provide a
'   simple text file (*.txt) browser.

'   Little or no error checking is implemented as it is designed purely as a
'   demo for the SplitterControls

'   Using the drive/folder/file controls was not the best choice as the
'   controls 'jump' to position and connont be resized smothly :-(

'[Author]
'   Richard Allsebrook  <RA>    RichardAllsebrook@easysoft.com

'[Declarations]
Option Explicit

Private Sub dir_Change()

'[Description]
'   Update the file list to refelect the path change

'[Code]

    fil.Path = dir.Path
    
End Sub

Private Sub drv_Change()

'[Description]
'   Update the folder list to reflect the drive change

'[Code]

    dir.Path = drv.Drive
    
End Sub

Private Sub fil_DblClick()

'[Description]
'   Load the double clicked file into the textbox

'[Notes]
'   There are much faster ways to read a file, but this is just a demo for the
'   EBSplitter so I'm not going to spend time optimizing!

'[Declarations]
Dim fh                      As Integer          'File handle of file we are viewing
Dim strLine                 As String           'Used to read the file

'[Code]

    'Open the file
    fh = FreeFile
    Open fil.Path & IIf(Right(fil.Path, 1) = "\", "", "\") & fil.FileName For Input As #fh
    
    If LOF(fh) > 32000 Then
        'textfile to big to display
        MsgBox "That text file is too large to display", vbInformation + vbOKOnly, "EBSplitter Test"
        
    Else
    
        Screen.MousePointer = vbHourglass
        
        'Load the file into the textbox
        txt.Text = ""
        
        Do While Not EOF(fh)
            Line Input #fh, strLine
            txt.Text = txt.Text & strLine & vbCrLf
        Loop
        Screen.MousePointer = vbNormal
        
    End If
    
    Close #fh
    
End Sub

Private Sub Form_Load()

'[Description]
'   Reposition the splitterbar (and force a redraw of its contained controls)

'[Code]

    ctlSplitterMain.SplitterPos = Me.Width / 3
    
End Sub

Private Sub Form_Resize()

'[Description]
'   Resize the main splitter control to fit the new form dimentions

'[Code]

    ctlSplitterMain.Move 0, 0, Me.Width - 120, Me.Height - 405
    
End Sub

Private Sub picExplorer_Resize()

'[Description]
'   Resize the drive box and Folder/File Splitter to match the new Explorer size

'[Code]

    With picExplorer
    
        If .Width > 45 And .Height - drv.Height > 0 Then
            drv.Width = .Width - 45
            ctlSplitterFolderFile.Move 0, drv.Height, .Width - 45, .Height - drv.Height
        End If
        
    End With
    
End Sub
