Attribute VB_Name = "DetailFinderModule"
'This module contains this program's core procedures.
Option Explicit

'Defines the Microsoft Windows API functions used by this program.
Public Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long

'This enumeration lists the supported highlighting modes.
Public Enum HighlightModesE
   BoxedHighlightMode = 0   'Highlight an image's sections using boxes.
   InvertedHightlightMode   'Highlight an image's sections by inverting their colors.
   ObscuredHighlightMode    'Highlight an image's sections by obscuring those that fall outside the selection.
End Enum

'This structure defines the results of an image analysis.
Public Type DetailsStr
   Levels() As Long        'Defines the detail levels for each image section.
   StepX As Long           'Defines the horizontal size of an image section.
   StepY As Long           'Defines the vertical size of an image section.
End Type

'This procedure combines the specified directory path and file name.
Public Function CombinePath(Path As String, FileName As String) As String
On Error GoTo ErrorTrap
Dim Combined As String

   Combined = Path
   If Not Right$(Combined, 1) = "\" Then Combined = Combined & "\"
   
EndRoutine:
   CombinePath = Combined & FileName
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure analyzes the specified image and returns the results.
Public Function GetDetails(Optional ImageBox As PictureBox = Nothing, Optional Precision As Long = 0, Optional XSteps As Long = Empty, Optional YSteps As Long = Empty) As DetailsStr
On Error GoTo ErrorTrap
Dim Blue1 As Long
Dim Blue2 As Long
Dim Color1 As Long
Dim Color2 As Long
Dim DetailLevel As Long
Dim Green1 As Long
Dim Green2 As Long
Dim HighestLevel As Long
Dim Red1 As Long
Dim Red2 As Long
Dim XSection As Long
Dim YSection As Long
Dim x As Long
Dim y As Long
Dim XDetailLevel As Long
Dim YDetailLevel As Long
Static Details As DetailsStr

   If Not Precision = 0 Then
      With Details
         HighestLevel = -1
         ReDim .Levels(0 To XSteps + 1, 0 To YSteps + 1) As Long
         .StepX = (ImageBox.ScaleWidth - 1) \ XSteps
         .StepY = (ImageBox.ScaleHeight - 1) \ YSteps
      End With
   
      With ImageBox
         .Enabled = False
      
         For x = 0 To .ScaleWidth - 1 Step Details.StepX
            For y = 0 To .ScaleHeight - 1 Step Details.StepY
               DetailLevel = 0
               For XSection = x To x + (Details.StepX - 1)
                  If XSection >= .ScaleWidth - 1 Then Exit For
                  For YSection = y To y + (Details.StepY - 1)
                     If YSection >= .ScaleHeight - 1 Then Exit For
                     Color1 = .Point(XSection, YSection)
                     Red1 = (Color1 And &HFF&)
                     Green1 = (Color1 And &HFF00&) / &H100&
                     Blue1 = (Color1 And &HFF0000) / &H10000
                     
                     Color2 = .Point(XSection + 1, YSection)
                     Red2 = (Color2 And &HFF&)
                     Green2 = (Color2 And &HFF00&) / &H100&
                     Blue2 = (Color2 And &HFF0000) / &H10000
                     XDetailLevel = Abs(Red1 - Red2) + Abs(Green1 - Green2) + Abs(Blue1 - Blue2)
                     
                     Color2 = .Point(XSection, YSection + 1)
                     Red2 = (Color2 And &HFF&)
                     Green2 = (Color2 And &HFF00&) / &H100&
                     Blue2 = (Color2 And &HFF0000) / &H10000
                     YDetailLevel = Abs(Red1 - Red2) + Abs(Green1 - Green2) + Abs(Blue1 - Blue2)
                     
                     DetailLevel = DetailLevel + ((XDetailLevel + YDetailLevel) / 2)
                  Next YSection
               Next XSection
               If DetailLevel > HighestLevel Then HighestLevel = DetailLevel
               Details.Levels(x \ Details.StepX, y \ Details.StepY) = DetailLevel
            Next y
            If DoEvents() = 0 Then Exit For
         Next x
      End With
      
      If Not HighestLevel = 0 Then
         With Details
            HighestLevel = HighestLevel / Precision
            For x = LBound(.Levels(), 1) To UBound(.Levels(), 1)
               For y = LBound(.Levels(), 2) To UBound(.Levels(), 2)
                  .Levels(x, y) = CLng(.Levels(x, y) / HighestLevel)
               Next y
            Next x
         End With
      End If
   
      ImageBox.Enabled = True
   End If

EndRoutine:
   GetDetails = Details
   Exit Function

ErrorTrap:
   HandleError
   Erase Details.Levels()
   Resume EndRoutine
End Function

'This procedure manages the currently loaded source image.
Public Function GetImage(Optional Path As String = vbNullString, Optional NewImage As Picture = Nothing) As Picture
On Error GoTo ErrorTrap
Static CurrentImage As Picture
   
   If Not NewImage Is Nothing Then
      Set CurrentImage = NewImage
   ElseIf Not Path = vbNullString Then
      Set CurrentImage = LoadPicture(Path)
   End If
   
EndRoutine:
   Set GetImage = CurrentImage
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure highlights the details in the specified image.
Public Sub HighlightDetails(ImageBox As PictureBox, Details As DetailsStr, LowerThreshold As Long, UpperThreshold As Long, Color As Long, Mode As HighlightModesE)
On Error GoTo ErrorTrap
Dim x As Long
Dim y As Long

   If Not SafeArrayGetDim(Details.Levels()) = 0 Then
      With ImageBox
         Set .Picture = GetImage()
         
         If Mode = InvertedHightlightMode Then
            .DrawMode = vbInvert
         Else
            .DrawMode = vbCopyPen
         End If
      
         For x = 0 To .ScaleWidth - 1 Step Details.StepX
            For y = 0 To .ScaleHeight - 1 Step Details.StepY
               If Details.Levels(x \ Details.StepX, y \ Details.StepY) >= LowerThreshold And Details.Levels(x \ Details.StepX, y \ Details.StepY) <= UpperThreshold Then
                  If Mode = BoxedHighlightMode Then
                     ImageBox.Line (x, y)-Step(Details.StepX, Details.StepY), Color, B
                  ElseIf Mode = InvertedHightlightMode Then
                     ImageBox.Line (x, y)-Step(Details.StepX - 1, Details.StepY - 1), , BF
                  End If
               Else
                  If Mode = ObscuredHighlightMode Then
                     ImageBox.Line (x, y)-Step(Details.StepX - 1, Details.StepY - 1), Color, BF
                  End If
               End If
            Next y
         Next x
      End With
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   On Error Resume Next
   MsgBox "Error: " & CStr(ErrorCode) & vbCr & Description, vbExclamation
End Sub

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   InterfaceWindow.Show
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns this program's information.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
   With App
      ProgramInformation = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & ", by: " & .CompanyName
   End With
EndRoutine:
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

