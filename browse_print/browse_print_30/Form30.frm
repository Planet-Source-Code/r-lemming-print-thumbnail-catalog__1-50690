VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form30 
   Caption         =   "Image Browser"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   917
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint30 
      Caption         =   "Print The 30 Currently Shown"
      Height          =   735
      Left            =   11040
      TabIndex        =   13
      Top             =   5160
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview Print"
      Height          =   735
      Left            =   11040
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   10440
      ScaleHeight     =   3795
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox picThumb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1152
      Left            =   12360
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   1536
   End
   Begin VB.PictureBox PicSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   11160
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   10440
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraSelection 
      Caption         =   "Image Browser"
      Height          =   7815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10095
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next 30"
         Height          =   735
         Left            =   1680
         TabIndex        =   12
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Previous 30"
         Height          =   735
         Left            =   240
         TabIndex        =   11
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoadPictures 
         Caption         =   "Create Thumbnails --->"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   6240
         Width           =   2415
      End
      Begin VB.DirListBox dlb 
         Height          =   5265
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.DriveListBox drv 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin MSComctlLib.ListView flb 
         Height          =   7215
         Left            =   3240
         TabIndex        =   1
         Top             =   480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   12726
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Items in List"
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblPreview 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Double click to preview"
         Height          =   195
         Left            =   6960
         TabIndex        =   4
         Top             =   840
         Width           =   1665
      End
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private lastPosition As Integer    'used in preview & print


' ++++++++++++++++++++++++++++++++ BROWSING - START ++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ BROWSING - START ++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ BROWSING - START ++++++++++++++++++++++++++++++++

' THE FOLLOWING CODE FOR BROWSING IS UNALTERED AND CAME FROM PSC:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=31026&lngWId=1
' ImageBrows502021212002.zip


Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260

Private Type FILETIME
       dwLowDateTime As Long
       dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
       dwFileAttributes As Long
       ftCreationTime As FILETIME
       ftLastAccessTime As FILETIME
       ftLastWriteTime As FILETIME
       nFileSizeHigh As Long
       nFileSizeLow As Long
       dwReserved0 As Long
       dwReserved1 As Long
       cFileName As String * MAX_PATH
       cAlternate As String * 14
End Type

Dim EnablePreview As Boolean
Dim Filename As String
Dim INIPath As String
Dim lstFilesFocus As Boolean

Dim flbList As New Collection


Private Sub cmdLoadPictures_Click()
  lastPosition = 1
  Load_The_Pics 1
End Sub


Private Sub cmdNext_Click()
   Load_The_Pics lastPosition
End Sub


Private Sub cmdPrevious_Click()
    Load_The_Pics lastPosition - 60
End Sub


Private Sub Load_The_Pics(startPosition As Integer) 'Sub to load the pics

    Dim i As Long
    Dim FN As String
    Dim hHeight As Double, hWidth As Double
      
    For i = flbList.Count To 1 Step -1
        flbList.Remove (i)
    Next
    
    flb.Icons = Nothing
    ImgList.ListImages.Clear
    
    flb.ListItems.Clear
    flb.Refresh
    
    GetFiles dlb.Path
    
    For i = flbList.Count To 1 Step -1
        FN = LCase$(Right$(flbList.Item(i), 3))
        If FN <> "jpg" And FN <> "bmp" And FN <> "cur" And FN <> "ico" Then
            flbList.Remove (i)
        End If
    Next
    
    Dim upperLimit As Integer
    upperLimit = 30
    If flbList.Count < 30 Then
       upperLimit = flbList.Count
    End If
    
    For i = 1 To upperLimit
    
        On Error Resume Next
        PicSrc.Picture = LoadPicture(flbList(startPosition))
        
        hWidth = PicSrc.Width
        hHeight = PicSrc.Height
        
        If hHeight > 76.8 Then
            hWidth = 76.8 * PicSrc.Width / PicSrc.Height
            hHeight = 76.8
        End If
        
        If hWidth > 102.4 Then
            hHeight = 102.4 * PicSrc.Height / PicSrc.Width
            hWidth = 102.4
        End If
        
        picThumb.PaintPicture PicSrc, (picThumb.Width - hWidth) / 2, (picThumb.Height - hHeight) / 2, hWidth, hHeight
        ImgList.ListImages.Add , , picThumb.Image
        If flb.Icons Is Nothing Then flb.Icons = ImgList
        flb.ListItems.Add , , GetFileName(flbList(startPosition)), i
        
        picThumb.Cls
        
        Caption = "GENERATING PREVIEWS  " & Format(Round(i / flbList.Count * 100, 2), "###.00") & "%"
        startPosition = startPosition + 1
    Next
    
    lastPosition = startPosition
    flb.Arrange = lvwAutoTop
    lblInfo.Caption = flb.ListItems.Count & " items in list"
    Caption = "Image Browser"
End Sub


Private Sub drv_Change()
    On Error GoTo Err
    dlb.Path = drv.Drive

Exit Sub
Err:
If Err.Number = 68 Then
    drv.Drive = "C:"
End If
End Sub

Private Sub GetFiles(Path As String)
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long, fPath As String, fName As String
   Dim colFiles As Collection
   Dim varFile As Variant
   
   fPath = AddBackslash(Path)
   fName = fPath & "*.*"
   Set colFiles = New Collection
   
   hFile = FindFirstFile(fName, WFD)
   If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
       colFiles.Add fPath & StripNulls(WFD.cFileName)
   End If
   
   While FindNextFile(hFile, WFD)
       If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
           colFiles.Add fPath & StripNulls(WFD.cFileName)
       End If
   Wend
   
   FindClose hFile
   
   For Each varFile In colFiles
       flbList.Add varFile
   Next
   Set colFiles = Nothing
End Sub

Private Function StripNulls(f As String) As String
   StripNulls = Left$(f, InStr(1, f, Chr$(0)) - 1)
End Function

Private Function AddBackslash(S As String) As String
   If Len(S) Then
      If Right$(S, 1) <> "\" Then
         AddBackslash = S & "\"
      Else
         AddBackslash = S
      End If
   Else
      AddBackslash = "\"
   End If
End Function

Private Function GetFileName(File As String) As String
    Dim i As Integer
    For i = Len(File) To 1 Step -1
        If Mid$(File, i, 1) = "\" Then
            i = i + 1
            Exit For
        End If
    Next
    
    GetFileName = Mid$(File, i)
End Function

Private Sub flb_DblClick()
    Dim Filename As String
    
    Filename = AddBackslash(dlb.Path)
    Filename = Filename & flb.SelectedItem
    
    ShellExecute Form30.hwnd, "", Filename, "", dlb.Path, 0
End Sub


' ++++++++++++++++++++++++++++++++ BROWSING - END ++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ BROWSING - END ++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ BROWSING - END ++++++++++++++++++++++++++++++++


' ++++++++++++++++++++++++++++++++ PREVIEW & PRINT - BEGIN +++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ PREVIEW & PRINT - BEGIN +++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ PREVIEW & PRINT - BEGIN +++++++++++++++++++++++

'THE FOLLOWING CODE FOR PREVIEW & PRINT ORIGINATED FROM (well worth looking up):
'http://support.microsoft.com/default.aspx?scid=http://support.microsoft.com:80/support/kb/articles/Q193/3/79.ASP&NoWebContent=1
'Microsoft Knowledge Base Article - 193379
'HOWTO: Print Preview in Visual Basic Applications


Private Sub cmdPreview_Click()   'preview
         Dim dRatio As Double
         dRatio = ScalePicPreviewToPrinterInches(Picture1)
         PrintRoutine30 lastPosition - 30, Picture1, dRatio
End Sub

Private Sub cmdPrint30_Click()   'print
         Printer.ScaleMode = vbInches
         PrintRoutine30 lastPosition - 30, Printer
         Printer.EndDoc
End Sub


Private Function ScalePicPreviewToPrinterInches(picPreview As PictureBox) As Double

         Dim Ratio As Double ' Ratio between Printer and Picture
         Dim LRGap As Double, TBGap As Double
         Dim HeightRatio As Double, WidthRatio As Double
         Dim PgWidth As Double, PgHeight As Double
         Dim smtemp As Long

         ' Get the physical page size in Inches:
         PgWidth = Printer.Width / 1440
         PgHeight = Printer.Height / 1440

         ' Find the size of the non-printable area on the printer to
         ' use to offset coordinates. These formulas assume the
         ' printable area is centered on the page:
         smtemp = Printer.ScaleMode
         Printer.ScaleMode = vbInches
         LRGap = (PgWidth - Printer.ScaleWidth) / 2
         TBGap = (PgHeight - Printer.ScaleHeight) / 2
         Printer.ScaleMode = smtemp

         ' Scale PictureBox to Printer's printable area in Inches:
         picPreview.ScaleMode = vbInches

         ' Compare the height and with ratios to determine the
         ' Ratio to use and how to size the picture box:
         HeightRatio = picPreview.ScaleHeight / PgHeight
         WidthRatio = picPreview.ScaleWidth / PgWidth

         If HeightRatio < WidthRatio Then
            Ratio = HeightRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = vbInches
            picPreview.Width = PgWidth * Ratio
            picPreview.Container.ScaleMode = smtemp
         Else
            Ratio = WidthRatio
            smtemp = picPreview.Container.ScaleMode
            picPreview.Container.ScaleMode = vbInches
            picPreview.Height = PgHeight * Ratio
            picPreview.Container.ScaleMode = smtemp
         End If

         ' Set default properties of picture box to match printer
         ' There are many that you could add here:
         picPreview.Scale (0, 0)-(PgWidth, PgHeight)
         picPreview.Font.Name = Printer.Font.Name
         picPreview.FontSize = Printer.FontSize * Ratio
         picPreview.ForeColor = Printer.ForeColor
         picPreview.Cls

         ScalePicPreviewToPrinterInches = Ratio
End Function
      
      
Private Sub PrintRoutine30(startPosition As Integer, objPrint As Object, Optional Ratio As Double = 1)
         ' All dimensions in inches:
         Dim xPosition As Double      'horizontal (or left) position of picture
         Dim yPosition As Double      'vertical (or top) position of picture
         xPosition = 0.5
         yPosition = 1
         Dim i As Integer
         
         Dim picWidth As Double       'picture width
         Dim picHeight As Double      'picture height
         picWidth = 1.3
         picHeight = 1
         
         Dim xSpacing As Double       'horizontal spacing bewtween pictures
         Dim ySpacing As Double       'vertical spacing between pictures
         xSpacing = 0.2
         ySpacing = 0.4
         
         Dim upperLimit As Integer
         upperLimit = 30
         If flbList.Count - startPosition < 30 Then
            upperLimit = flbList.Count - startPosition + 1
         End If

     For i = 1 To upperLimit

         ' Print some graphics to the control object
         PicSrc.Picture = LoadPicture(flbList(startPosition))
         'object.PaintPicture picture, x1, y1, width1, height1, x2, y2, width2, height2, opcode   '<-- general format
         objPrint.PaintPicture PicSrc.Picture, xPosition, yPosition, picWidth, picHeight

         ' Print a the filename
         With objPrint
            .Font.Name = "Arial"
            .CurrentX = xPosition
            .CurrentY = yPosition + (picHeight + (ySpacing / 2))  'filename goes between the pictures, so divide by 2
            .FontSize = 7 * Ratio
            objPrint.Print flb.ListItems(i)                       'the filename
         End With
         
         xPosition = xPosition + (picWidth + xSpacing)       'next picture moves in the x direction

         If xPosition >= 8 Then                              'if xPosition is greater than 8 in., then start a new row
            xPosition = 0.5                                  'new row so x starts a the beginning
            yPosition = yPosition + (picHeight + ySpacing)   'y moves down one row
         End If
         
         startPosition = startPosition + 1
         
     Next i

End Sub
      

' ++++++++++++++++++++++++++++++++ PREVIEW & PRINT - END +++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ PREVIEW & PRINT - END +++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++ PREVIEW & PRINT - END +++++++++++++++++++++++




