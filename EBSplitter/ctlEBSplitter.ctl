VERSION 5.00
Begin VB.UserControl ctlEBSplitter 
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ControlContainer=   -1  'True
   ScaleHeight     =   855
   ScaleWidth      =   1935
   ToolboxBitmap   =   "ctlEBSplitter.ctx":0000
   Begin VB.PictureBox picSplitter 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   960
      MousePointer    =   9  'Size W E
      ScaleHeight     =   855
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   0
      Width           =   75
   End
End
Attribute VB_Name = "ctlEBSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'[Description]
'   EBSplitter
'   A VB6 splitter control

'[Author]
'   Richard Allsebrook  <RA>    RichardAllsebrook@earlybirdmarketing.com

'[History]
'   22/06/2001  V1.0.0
'   Initial development

'[Declarations]
Option Explicit

'Property Storage
Private zOrientation        As EBSplitterOrientation    'Current orientation of splitterbar
Private lngSplitterWidth    As Long         'Width of splitterbar (in twips)

'Enumerations
Public Enum EBSplitterOrientation
    EBHorizontal
    EBVertical
End Enum

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'[Description]
'   This is the workhorse routine for the control.  If caluclates the new position of the
'   splitter bar, move it and forces a resize of the contained controls

'[Notes]
'   (1) We check the new position is different to the last position as some
'   controls develope an annouying flicker when they are resized to the same
'   dimentions repeatedly


'[Declarations]
Dim lngPos                  As Long         'New position of splitterbar
Static lngLastHPos          As Long         'Last Horizontal position of splitterbar
Static lngLastVPos          As Long         'Last Vertical position of splitterbar

'[Code]

    If Button = vbLeftButton Then
        'Only move the bar if the left button is pressed
        
        If zOrientation = EBHorizontal Then
            'Move vertically
            lngPos = ((picSplitter.Top + Y - lngSplitterWidth / 2) \ 15) * 15
            
            'Constrain splitter to current control size
            If lngPos < 120 Then
                'Attempted to move past start
                lngPos = 120
                
            ElseIf lngPos > UserControl.Height - lngSplitterWidth - 120 Then
            
                'Attempted to move past end
                lngPos = UserControl.Height - lngSplitterWidth - 120
            End If
            
            If lngPos <> lngLastVPos Then
                'See note (1)
                lngLastVPos = lngPos
                picSplitter.Move 0, lngPos
                PropertyChanged "SplitterPos"
                ResizeControls
            End If
            
        Else
            'Move horizontally
            lngPos = ((picSplitter.Left + X - lngSplitterWidth / 2) \ 15) * 15
            
            'Constrain splitter to current control size
            If lngPos < 120 Then
                'Attempted to move past start
                lngPos = 120
                
            ElseIf lngPos > UserControl.Width - lngSplitterWidth - 120 Then
            
                'Attempted to move past end
                lngPos = UserControl.Width - lngSplitterWidth - 120
            End If
            
            If lngPos <> lngLastHPos Then
                'See Note (1)
                lngLastHPos = lngPos
                picSplitter.Move lngPos, 0
                PropertyChanged "SplitterPos"
                
                ResizeControls
            End If
            
        End If
        
    End If
    
End Sub

Private Sub UserControl_InitProperties()

'[Description]
'   Initialise the control's properties

'[Code]

    lngSplitterWidth = 75
    Orientation = EBVertical
    picSplitter.Left = UserControl.Width / 2
    
End Sub

Public Property Get Orientation() As EBSplitterOrientation

'[Description]
'   Return the current splitterbar orientation

'[Code]

    Orientation = zOrientation
    
End Property

Public Property Let Orientation(NewValue As EBSplitterOrientation)

'[Description]
'   Sets the splitterbar orientation and redraw the control accordingly

'[Code]

    zOrientation = NewValue
    
    'Resize/move the splitterbar to fit the new orientation
    If zOrientation = EBHorizontal Then
        picSplitter.Move 0, UserControl.Height / 2, UserControl.Width, lngSplitterWidth
        picSplitter.MousePointer = vbSizeNS
    Else
        picSplitter.Move UserControl.Width / 2, 0, lngSplitterWidth, UserControl.Height
        picSplitter.MousePointer = vbSizeWE
    End If
    
    ResizeControls
    
    PropertyChanged "Orientation"
    PropertyChanged "SplitterPos"
    
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'[Description]
'   Restore the control's properties from its PropBag

'[Code]

    With PropBag
        lngSplitterWidth = .ReadProperty("SplitterWidth", 75)
        Orientation = .ReadProperty("Orientation", EBVertical)
        
        If zOrientation = EBHorizontal Then
            picSplitter.Top = .ReadProperty("SplitterPos", UserControl.Height / 2)
        Else
            picSplitter.Left = .ReadProperty("SplitterPos", UserControl.Width / 2)
        End If
        
    End With
    
    ResizeControls
    
End Sub

Private Sub UserControl_Resize()

'[Description]
'   Force a resize of the constituant and contained controls

'[Code]

    ResizeControls
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'[Description]
'   Store the controls current property settings in its PropBag

'[Code]

    With PropBag
        .WriteProperty "SplitterWidth", lngSplitterWidth, 75
        .WriteProperty "Orientation", zOrientation, EBVertical
        
        If zOrientation = EBHorizontal Then
            .WriteProperty "SplitterPos", picSplitter.Top, UserControl.Height / 2
        Else
            .WriteProperty "SplitterPos", picSplitter.Left, UserControl.Width / 2
        End If
        
    End With
    
End Sub

Private Function ResizeControls()

'[Description]
'   Resize the contained controls (1 or 2) to match the new control size and
'   SplitterPos

'[Declarations]
Dim lngPos                  As Long         'Current Splitter pos

'[Code]

    With UserControl
    
        If zOrientation = EBHorizontal Then
            'Resize splitterbar
            picSplitter.Left = 0
            picSplitter.Width = .Width
            picSplitter.Height = lngSplitterWidth
            
            'Resize contained controls
            If .ContainedControls.Count = 1 Then
                .ContainedControls(0).Move 0, 0, .Width, picSplitter.Top + 15
                
            ElseIf .ContainedControls.Count = 2 Then
            
                lngPos = picSplitter.Top + lngSplitterWidth - 15
                .ContainedControls(1).Move 0, 0, .Width, picSplitter.Top + 15
                
                If .Height - lngPos > 0 Then
                    .ContainedControls(0).Move 0, lngPos, .Width, .Height - lngPos
                End If
                
            End If
            
        Else
            'Resize splitterbar
            picSplitter.Top = 0
            picSplitter.Height = .Height
            picSplitter.Width = lngSplitterWidth
            
            'Resize contained controls
            If .ContainedControls.Count = 1 Then
                .ContainedControls(0).Move 0, 0, picSplitter.Left + 15, .Height
                
            ElseIf .ContainedControls.Count = 2 Then
            
                lngPos = picSplitter.Left + lngSplitterWidth - 15
                .ContainedControls(1).Move 0, 0, picSplitter.Left + 15, .Height
                
                If .Width - lngPos > 0 Then
                    .ContainedControls(0).Move lngPos, 0, .Width - lngPos, .Height
                End If
                
            End If
            
        End If
        
    End With

End Function

Public Property Get SplitterWidth() As Long

'[Description]
'   Return the current SplitterWidth

'[Code]

    SplitterWidth = lngSplitterWidth
    
End Property

Public Property Let SplitterWidth(NewValue As Long)

'[Description]
'   Set the current SplitterWidth and redraw contained controls

'[Code]

    lngSplitterWidth = NewValue
    ResizeControls
    
    PropertyChanged "SplitterWidth"
    
End Property

Public Property Get SplitterPos() As Long

'[Description]
'   Return the current Splitterbar position (dependant on orientation)

'[Code]

    If zOrientation = EBHorizontal Then
        SplitterPos = picSplitter.Top
    Else
        SplitterPos = picSplitter.Left
    End If
    
End Property

Public Property Let SplitterPos(NewValue As Long)

'[Description]
'   Set the new splitterpos

'[Code]

    If zOrientation = EBHorizontal Then
    
        'Contrain splitter to within control vertical boundaries
        If NewValue < 120 Then
            NewValue = 120
            
        ElseIf NewValue > UserControl.Height - lngSplitterWidth - 120 Then
        
            NewValue = UserControl.Height - lngSplitterWidth - 120
        End If
        
        picSplitter.Top = NewValue
    Else
    
        'Contrain splitter to within control horizontal boundaries
        If NewValue < 120 Then
            NewValue = 120
            
        ElseIf NewValue > UserControl.Width - lngSplitterWidth - 120 Then
        
            NewValue = UserControl.Width - lngSplitterWidth - 120
        End If
        
        picSplitter.Left = NewValue
    End If
    
    ResizeControls
    PropertyChanged "SplitterPos"
    
End Property
