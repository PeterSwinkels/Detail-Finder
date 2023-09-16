VERSION 5.00
Begin VB.Form InterfaceWindow 
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13155
   ClipControls    =   0   'False
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   35.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   109.625
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ImagePanel 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   3855
      Left            =   2880
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   240
      Width           =   6495
      Begin VB.PictureBox CornerBox 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   255
         Left            =   6240
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3600
         Width           =   255
      End
      Begin VB.HScrollBar HorizontalScrollbar 
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   3600
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.VScrollBar VerticalScrollbar 
         Height          =   3615
         Left            =   6240
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox ImageBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   975
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame ImageFrame 
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox GrabFromClipboardBox 
         Alignment       =   1  'Right Justify
         Caption         =   "Grab from &clipboard:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   3
         ToolTipText     =   "Select to enable the automatic analysis of any image copied to the clipboard."
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Timer ImageGrabber 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   120
         Top             =   3600
      End
      Begin VB.FileListBox FileListBox 
         Height          =   1845
         Hidden          =   -1  'True
         Left            =   120
         Pattern         =   "*.bmp;*.cur;*.emf;*.gif;*.ico;*.jpg;*.jpeg;*.rle;*.wmf"
         System          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Double click to load the selected image."
         Top             =   240
         Width           =   2295
      End
      Begin VB.DirListBox DirectoryListBox 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Select the directory containing the image to analyze here."
         Top             =   2160
         Width           =   2295
      End
      Begin VB.DriveListBox DriveListBox 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Select the drive containing the image to analyze here."
         Top             =   3480
         Width           =   2415
      End
   End
   Begin VB.Frame AnalysisFrame 
      Caption         =   "Analysis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   2655
      Begin VB.CommandButton AnalyzeButton 
         Caption         =   "&Analyze"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Click here to start the analysis of the detail levels in the current image."
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox PrecisionBox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Specify the precision used to measure detail levels here."
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox YGridSizeBox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Specify the number of vertical image sections here."
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox XGridSizeBox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Specify the number of horizontal image sections here."
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Precision:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   ","
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   22
         Top             =   600
         Width           =   135
      End
   End
   Begin VB.Frame HighlightFrame 
      Caption         =   "Highlight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   6120
      Width           =   2655
      Begin VB.TextBox UpperThresholdBox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         ToolTipText     =   "Specify the highest detail level to be highlighted here."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox LowerThresholdBox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   8
         ToolTipText     =   "Specify the lowest detail level to be highlighted here."
         Top             =   240
         Width           =   495
      End
      Begin VB.HScrollBar HighlightColorBar 
         Height          =   255
         Left            =   840
         Max             =   15
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.PictureBox HighlightColorBox 
         ClipControls    =   0   'False
         Height          =   255
         Left            =   840
         ScaleHeight     =   0.813
         ScaleMode       =   4  'Character
         ScaleWidth      =   12.625
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "The color used to highlight the image sections with the specified detail levels."
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton HighlightModeBox 
         Caption         =   "Highlight: &boxes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Highlight image sections using boxes."
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton HighlightModeBox 
         Caption         =   "Highlight: &inverted."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Highlight image sections by inverting their colors."
         Top             =   1560
         Width           =   2055
      End
      Begin VB.OptionButton HighlightModeBox 
         Caption         =   "Highlight: &obscured."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Highlight image sections by obscuring those that fall outside the specified range."
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Levels:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   19
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface.
Option Explicit

'This procedure returns the selected highlight mode based upon which highlight mode box is on.
Private Function GetHighlightMode() As HighlightModesE
On Error GoTo ErrorTrap
   Dim Index As Integer
   Dim Mode As HighlightModesE
   
   Mode = BoxedHighlightMode
   For Index = HighlightModeBox.LBound() To HighlightModeBox.UBound()
      If HighlightModeBox(Index).Value = True Then
         Mode = CLng(Index)
         Exit For
      End If
   Next Index
EndRoutine:
   GetHighlightMode = Mode
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure determines whether the image scroll bars are visible or not.
Private Sub ToggleScrollBars()
On Error GoTo ErrorTrap
   HorizontalScrollbar.Visible = (ImageBox.ScaleWidth > ImagePanel.ScaleWidth)
   VerticalScrollbar.Visible = (ImageBox.ScaleHeight > ImagePanel.ScaleHeight)
   CornerBox.Visible = (HorizontalScrollbar.Visible Or VerticalScrollbar.Visible)
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to analyze the current image.
Private Sub AnalyzeButton_Click()
On Error GoTo ErrorTrap
   Screen.MousePointer = vbHourglass
   Set ImageBox.Picture = GetImage()
   HighlightDetails ImageBox, GetDetails(ImageBox, CLng(Val(PrecisionBox.Text)), CLng(Val(XGridSizeBox.Text)), CLng(Val(YGridSizeBox.Text))), CLng(Val(LowerThresholdBox.Text)), CLng(Val(UpperThresholdBox.Text)), QBColor(HighlightColorBar.Value), GetHighlightMode()
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure updates the file list when the selected directory changes.
Private Sub DirectoryListBox_Change()
On Error GoTo ErrorTrap
   FileListBox.Path = DirectoryListBox.Path
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure updates the directory list when the selected drive changes.
Private Sub DriveListBox_Change()
On Error GoTo ErrorTrap
   DirectoryListBox.Path = DriveListBox.Drive
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to load the selected image file.
Private Sub FileListBox_DblClick()
On Error GoTo ErrorTrap
   HorizontalScrollbar.Value = 0
   VerticalScrollbar.Value = 0
   Set ImageBox.Picture = GetImage(CombinePath(DirectoryListBox.Path, FileListBox.List(FileListBox.ListIndex)))
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Caption = ProgramInformation()
   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
   
   ImagePanel.Width = (Me.ScaleWidth - 3) - ImageFrame.Width
   ImagePanel.Height = Me.ScaleHeight - 2
   ImageBox.Width = ImagePanel.ScaleWidth
   ImageBox.Height = ImagePanel.ScaleHeight
   
   DriveListBox.Drive = Left$(CurDir$, InStr(CurDir$, ":"))
   DirectoryListBox.Path = CurDir$()
   FileListBox.ToolTipText = FileListBox.ToolTipText & " Supported types:" & Replace(FileListBox.Pattern, "*", " ")
   
   HighlightColorBar.Value = 10
   HighlightModeBox(BoxedHighlightMode).Value = True
   LowerThresholdBox.Text = "5"
   PrecisionBox.Text = "10"
   UpperThresholdBox.Text = "10"
   XGridSizeBox.Text = "10"
   YGridSizeBox.Text = "10"
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure adjusts this window's objects to its new size.
Private Sub Form_Resize()
On Error Resume Next
   ImagePanel.Width = (Me.ScaleWidth - 3) - ImageFrame.Width
   ImagePanel.Height = Me.ScaleHeight - 2
   
   CornerBox.Left = ImagePanel.ScaleWidth - CornerBox.ScaleWidth
   CornerBox.Top = ImagePanel.ScaleHeight - CornerBox.ScaleHeight
   HorizontalScrollbar.Max = ImageBox.ScaleWidth - (ImagePanel.ScaleWidth - VerticalScrollbar.Width)
   HorizontalScrollbar.Top = ImagePanel.ScaleHeight - HorizontalScrollbar.Height
   HorizontalScrollbar.Width = ImagePanel.ScaleWidth - VerticalScrollbar.Width
   VerticalScrollbar.Max = ImageBox.ScaleHeight - (ImagePanel.ScaleHeight - HorizontalScrollbar.Height)
   VerticalScrollbar.Left = ImagePanel.ScaleWidth - VerticalScrollbar.Width
   VerticalScrollbar.Height = ImagePanel.ScaleHeight - HorizontalScrollbar.Height
   
   ToggleScrollBars
End Sub

'This procedure enables/disables the image grabber.
Private Sub GrabFromClipboardBox_Click()
On Error GoTo ErrorTrap
   ImageGrabber.Enabled = (GrabFromClipboardBox.Value = vbChecked)
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to highlight the image sections using the selected color.
Private Sub HighlightColorBar_Change()
On Error GoTo ErrorTrap
   HighlightDetails ImageBox, GetDetails(), CLng(Val(LowerThresholdBox.Text)), CLng(Val(UpperThresholdBox.Text)), QBColor(HighlightColorBar.Value), GetHighlightMode()
   HighlightColorBox.BackColor = QBColor(HighlightColorBar.Value)
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure sets the highlight mode.
Private Sub HighlightModeBox_Click(Index As Integer)
On Error GoTo ErrorTrap
   HighlightDetails ImageBox, GetDetails(), CLng(Val(LowerThresholdBox.Text)), CLng(Val(UpperThresholdBox.Text)), QBColor(HighlightColorBar.Value), GetHighlightMode()
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure horizontally scrolls the current image.
Private Sub HorizontalScrollbar_Change()
On Error GoTo ErrorTrap
   ImageBox.Left = -HorizontalScrollbar.Value
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the information for the selected image section.
Private Sub ImageBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
Dim Details As DetailsStr

   Me.Caption = ProgramInformation()
   
   Details = GetDetails()
   With Details
      If Not SafeArrayGetDim(GetDetails.Levels()) = 0 Then
         Me.Caption = Me.Caption & " - " & " Detail Level: " & CStr(.Levels(X \ .StepX, Y \ .StepY))
      End If
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to load the file dropped into the image box.
Private Sub ImageBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
   If Data.Files.Count > 0 Then Set ImageBox.Picture = GetImage(Data.Files.Item(1))
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure updates the tooltip specifying the current image's size.
Private Sub ImageBox_Resize()
On Error GoTo ErrorTrap
   ImageBox.ToolTipText = "Size: " & CStr(ImageBox.ScaleWidth) & " x " & CStr(ImageBox.ScaleHeight)
   HorizontalScrollbar.Max = ImageBox.ScaleWidth - (ImagePanel.ScaleWidth - VerticalScrollbar.Width)
   VerticalScrollbar.Max = ImageBox.ScaleHeight - (ImagePanel.ScaleHeight - HorizontalScrollbar.Height)
   
   ToggleScrollBars
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure waits for an image to be copied to the clipboard and sets it as the current image.
Private Sub ImageGrabber_Timer()
On Error GoTo ErrorTrap
   If Clipboard.GetFormat(vbCFBitmap) Then
      GetImage , NewImage:=Clipboard.GetData(vbCFBitmap)
      Clipboard.Clear
      Set ImageBox = GetImage()
      HighlightDetails ImageBox, GetDetails(ImageBox, CLng(Val(PrecisionBox.Text)), CLng(Val(XGridSizeBox.Text)), CLng(Val(YGridSizeBox.Text))), CLng(Val(LowerThresholdBox.Text)), CLng(Val(UpperThresholdBox.Text)), QBColor(HighlightColorBar.Value), GetHighlightMode()
      Beep
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to highlight the image sections with the specified detail level.
Private Sub LowerThresholdBox_LostFocus()
On Error GoTo ErrorTrap
   HighlightDetails ImageBox, GetDetails(), CLng(Val(LowerThresholdBox.Text)), CLng(Val(UpperThresholdBox.Text)), QBColor(HighlightColorBar.Value), GetHighlightMode()
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub
'This procedure gives the command to highlight the image sections with the specified detail level.
Private Sub UpperThresholdBox_LostFocus()
On Error GoTo ErrorTrap
   HighlightDetails ImageBox, GetDetails(), CLng(Val(LowerThresholdBox.Text)), CLng(Val(UpperThresholdBox.Text)), QBColor(HighlightColorBar.Value), GetHighlightMode()
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure vertically scrolls the current image.
Private Sub VerticalScrollbar_Change()
On Error GoTo ErrorTrap
   ImageBox.Top = -VerticalScrollbar.Value
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


