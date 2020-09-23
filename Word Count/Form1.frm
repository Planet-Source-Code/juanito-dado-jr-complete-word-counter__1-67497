VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Word Counter"
   ClientHeight    =   8805
   ClientLeft      =   4605
   ClientTop       =   3615
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   14843
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Setup Page"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Dir1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "File1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ImageList1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Report Page"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Option1(4)"
      Tab(1).Control(1)=   "Option1(3)"
      Tab(1).Control(2)=   "Option1(2)"
      Tab(1).Control(3)=   "Option1(1)"
      Tab(1).Control(4)=   "cmdCopy"
      Tab(1).Control(5)=   "MSFlexGrid1"
      Tab(1).ControlCount=   6
      Begin VB.OptionButton Option1 
         Caption         =   "Words"
         Height          =   255
         Index           =   4
         Left            =   -67680
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Char W/Out Spaces"
         Height          =   255
         Index           =   3
         Left            =   -69600
         TabIndex        =   17
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Linecount"
         Height          =   255
         Index           =   2
         Left            =   -70800
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Char With Spaces"
         Height          =   255
         Index           =   1
         Left            =   -72480
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy to Clipboard"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7335
         Left            =   -74760
         TabIndex        =   13
         Top             =   840
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   12938
         _Version        =   393216
         Cols            =   6
         BackColorBkg    =   -2147483624
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10680
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0038
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":06DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0A2E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   11160
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.DirListBox Dir1 
         Height          =   7740
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   3420
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   4080
         Width           =   7335
         Begin VB.CommandButton cmdAddFile 
            Caption         =   "Add File"
            Height          =   375
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Add single item on the list"
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdRmvAll 
            Caption         =   "Remove All"
            Height          =   375
            Left            =   3480
            TabIndex        =   6
            ToolTipText     =   "Remove all files from the list"
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add All"
            Height          =   375
            Left            =   1080
            TabIndex        =   5
            ToolTipText     =   "Add all items on the list"
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdWordCount 
            Caption         =   "Start Counting"
            Default         =   -1  'True
            Height          =   375
            Left            =   5640
            TabIndex        =   4
            ToolTipText     =   "Start Word Count!!!"
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmdRmvFile 
            Caption         =   "Remove File"
            Height          =   375
            Left            =   2280
            TabIndex        =   3
            ToolTipText     =   "remove file from the list"
            Top             =   0
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3615
         Left            =   3840
         TabIndex        =   9
         Top             =   4560
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   14111
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3480
         Left            =   3840
         TabIndex        =   10
         ToolTipText     =   "Double click OR Drag the File"
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6138
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8550
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   10320
      TabIndex        =   12
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
'oooooooooooooooo Word Counter by Juanito Dado Jr oooooooooooooooooooooo
'ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo


Option Explicit
'progress bar in status bar
Private Declare Function SetParent Lib "user32" _
        (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Dim fs As FileSystemObject
Dim cnt As Long, i As Long, x As Long
Dim dirPath As String

Private Sub showfileinfo(fileSpec As String, x As Integer)
    Dim f
    Dim listitem1 As ListItem
    
    Set fs = New FileSystemObject
    Set f = fs.GetFile(fileSpec)
   
   'include path to the listview
    'replace \\ with \ if the  path is on c:\
    dirPath = Replace(Dir1.Path & "\" & fs.GetFileName(fileSpec), "\\", "\")
    
    Set listitem1 = ListView1.ListItems.Add()
    
    ' FileName
    listitem1.Text = dirPath
    ' Date Created
    ListView1.ListItems(x + 1).ListSubItems.Add , , Format(f.DateCreated, "mm/dd/yyyy")
    ' Date last Modified
    ListView1.ListItems(x + 1).ListSubItems.Add , , Format(f.DateLastModified, "mm/dd/yyyy")
    ' File Type
    ListView1.ListItems(x + 1).ListSubItems.Add , , f.Type
End Sub

Private Sub cmdCopy_Click()
    'copy to clipboard
    Call CopySelected
End Sub

Private Sub cmdRmvFile_Click()
    'if list is empty then exit
    If ListView2.ListItems.Count = 0 Then
        MsgBox "List Empty", vbInformation, "Word Counter"
        Exit Sub
    Else
        ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
    End If
End Sub

Private Sub cmdWordCount_Click()
Dim wordObject As Word.Application
Dim charWithSpace As Long, charNoSpace As Long, Words As Long
Dim x As Long
Dim timeEnd As String, timeStart As String
Dim lineCount As Double
Dim listy As ListItem
Dim strVar() As String
Dim strFile As String
    

    'set flexgrid to number of items on listview
    MSFlexGrid1.Rows = ListView2.ListItems.Count + 1
    'for selected column color
   
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No files available to Count", vbInformation, "Word Counter"
        Exit Sub
    Else
        ProgressBar1.Min = 0
        ProgressBar1.Max = ListView2.ListItems.Count
        timeStart = Format(Time, "hh:nn:ss")
        
        'set wordObject
        Set wordObject = New Word.Application
        
        For x = 1 To ListView2.ListItems.Count
            
            'set invalid line count to exe files
            If InStr(1, ListView2.ListItems.Item(x).Text, "exe") > 0 Then
                SSTab1.Tab = 1
                With MSFlexGrid1
                    'flexgrid alignment
                    .ColAlignment(1) = flexAlignLeftCenter
                    .TextMatrix(x, 1) = UseInStrRev(ListView2.ListItems.Item(x).Text)
                    .TextMatrix(x, 2) = "Invalid"
                    .TextMatrix(x, 3) = "Invalid"
                    .TextMatrix(x, 4) = "Invalid"
                    .TextMatrix(x, 5) = "Invalid"
                End With
            Else
                SSTab1.Tab = 1
                ProgressBar1.Value = x
                wordObject.Documents.Open ListView2.ListItems.Item(x).Text
                'compute for char with space
                charWithSpace = wordObject.ActiveDocument.Content.ComputeStatistics(wdStatisticCharactersWithSpaces)
                'compute for char without space
                charNoSpace = wordObject.ActiveDocument.Content.ComputeStatistics(wdStatisticCharacters)
                'compute for words
                Words = wordObject.ActiveDocument.Content.ComputeStatistics(wdStatisticWords)
                'linecount divided by 65
                lineCount = charWithSpace / 65
                'flexgrid alignment
                With MSFlexGrid1
                    .ColAlignment(1) = flexAlignLeftCenter
                    'populate msflexgrid with data
                    .TextMatrix(x, 1) = UseInStrRev(ListView2.ListItems.Item(x).Text)
                    .TextMatrix(x, 2) = CStr(charWithSpace)
                    .TextMatrix(x, 3) = CStr(Round(lineCount, 2))
                    .TextMatrix(x, 4) = CStr(charNoSpace)
                    .TextMatrix(x, 5) = CStr(Words)
                End With
            End If
        Next x

        wordObject.Quit
        Set wordObject = Nothing
        
        timeEnd = Format(Time, "hh:nn:ss")
        MsgBox "Duration: " & (DateDiff("s", timeStart, timeEnd)) & " second(s)", , "Word Counter"
        ProgressBar1.Value = 0
        Call colBackColor(MSFlexGrid1, 2, RGB(255, 255, 163))
        Option1(1) = True
    End If
End Sub

'for extracting filenames only
Private Function UseInStrRev(ByVal strIn As String) As String
Dim intPos As Integer

    intPos = InStrRev(strIn, "\") + 1
    UseInStrRev = Mid(strIn, intPos)
End Function


Private Sub cmdAddFile_Click()
    'if nothing selected
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Select items on the list above.", vbInformation, "Word Counter"
    Else
        Call AddToListview
    End If
End Sub

Private Sub cmdAddAll_Click()
Dim y As Long
    
    For y = 1 To ListView1.ListItems.Count
        'listview icon purposes
        'this can be done by shell and API but i dunno how to do it. ^ ^
        If InStr(1, ListView1.ListItems.Item(y).Text, "doc") > 0 Or InStr(1, ListView1.ListItems.Item(y).Text, "DOC") > 0 _
        Or InStr(1, ListView1.ListItems.Item(y).Text, "rtf") > 0 Or InStr(1, ListView1.ListItems.Item(y).Text, "RTF") > 0 Then
            ListView2.ListItems.Add , , ListView1.ListItems.Item(y).Text, , 1
        ElseIf InStr(1, ListView1.ListItems.Item(y).Text, "xls") > 0 Or InStr(1, ListView1.ListItems.Item(y).Text, "XLS") > 0 Then
            ListView2.ListItems.Add , , ListView1.ListItems.Item(y).Text, , 2
        ElseIf InStr(1, ListView1.ListItems.Item(y).Text, "txt") > 0 Or InStr(1, ListView1.ListItems.Item(y).Text, "TXT") > 0 Then
            ListView2.ListItems.Add , , ListView1.ListItems.Item(y).Text, , 3
        Else
            ListView2.ListItems.Add , , ListView1.ListItems.Item(y).Text, , 4
        End If
    Next y
End Sub

Private Sub cmdRmvAll_Click()
    ListView2.ListItems.Clear
    MSFlexGrid1.Clear
    'for columheaders title
    MSFlexGrid1.FormatString = "|Filename|Char With Spaces|Linecount|Char W/Out Spaces|Words"
End Sub



Private Sub Dir1_Change()
    Dim x As Integer
    
    ListView1.ListItems().Clear
    File1.Path = Dir1.Path
    Screen.MousePointer = vbHourglass
    For x = 0 To File1.ListCount - 1
        showfileinfo File1.Path & "/" & File1.List(x), x
    Next x
    Screen.MousePointer = vbDefault
    'remove default selection on listview
    Call removeSelection
End Sub

Private Sub removeSelection()
'remove default selection on the listview
    With ListView1
        For cnt = 1 To .ListItems.Count
        .ListItems(cnt).Selected = False
        Next
        Set .SelectedItem = Nothing
    End With
End Sub
    

Private Sub Form_Load()
    'attaching the progress bar to statbar
    'setparent
    SetParent ProgressBar1.hWnd, StatusBar1.hWnd
    ProgressBar1.Top = 55
    'position
    ProgressBar1.Left = StatusBar1.Panels(1).Width + 60
    'size
    ProgressBar1.Width = StatusBar1.Panels(2).Width - 60
    ProgressBar1.Height = StatusBar1.Height - 90
    
    
    Dir1.Path = "c:\"
    Dim x As Integer
    
    ListView1.ListItems().Clear
    
    Screen.MousePointer = vbHourglass
    For x = 0 To File1.ListCount - 1
        showfileinfo File1.Path & "/" & File1.List(x), x
    Next x
    Screen.MousePointer = vbDefault
    
    ListView1.View = lvwReport
    ' Set up headers for listView
    Dim colHeader As ColumnHeader
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "Name"
    colHeader.Width = 5000
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "Date Created"
    colHeader.Width = 1500
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "Date Modified"
    colHeader.Width = 2000
    Set colHeader = ListView1.ColumnHeaders.Add()
    colHeader.Text = "File Type"
    colHeader.Width = 2000
    
    With MSFlexGrid1
    'msflexgrid formats
        .FormatString = "|Filename|Char With Spaces|Linecount|Char W/Out Spaces|Words"
        .ColAlignment(1) = flexAlignLeftCenter
    
    'msflexgrid size
        .ColWidth(0) = 400
        .ColWidth(1) = 6500
        .ColWidth(2) = 1500
        .ColWidth(4) = 1600
        .ColWidth(5) = 1000
    End With
    'remove default selection on listview
    Call removeSelection
End Sub

Private Sub Form_Resize()
'autoscale objects in side the form
On Error Resume Next   ' this is needed because when the user resize it to minimum it'll send an error
     ProgressBar1.Left = StatusBar1.Panels(1).Width + 60
     Dir1.Height = Me.Height - 1800
     ListView1.Height = Frame1.Top - 550
     ListView1.Width = Me.Width - 4200
     ListView2.Width = Me.Width - 4200
     ListView2.Height = Me.Height - 5900
     Frame1.Top = ListView2.Top - 450
     SSTab1.Height = Me.Height - 1200
     SSTab1.Width = Me.Width - 100
     MSFlexGrid1.Width = Me.Width - 550
     MSFlexGrid1.Height = Me.Height - 2200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fs = Nothing
End Sub

Private Sub AddToListview()
Dim z As Long
    For z = 1 To ListView2.ListItems.Count
            If ListView2.ListItems.Item(z).Text = ListView1.SelectedItem.Text Then
                MsgBox "Already added to the list.", vbInformation, "Word Counter"
                Exit Sub
            End If
    Next z
    
    If InStr(1, ListView1.SelectedItem.Text, "doc") > 0 Or InStr(1, ListView1.SelectedItem.Text, "DOC") > 0 _
    Or InStr(1, ListView1.SelectedItem.Text, "rtf") > 0 Or InStr(1, ListView1.SelectedItem.Text, "RTF") > 0 Then
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 1
    ElseIf InStr(1, ListView1.SelectedItem.Text, "xls") > 0 Or InStr(1, ListView1.SelectedItem.Text, "XLS") > 0 Then
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 2
    ElseIf InStr(1, ListView1.SelectedItem.Text, "txt") > 0 Or InStr(1, ListView1.SelectedItem.Text, "TXT") > 0 Then
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 3
    Else
        ListView2.ListItems.Add , , ListView1.SelectedItem.Text, , 4
    End If
End Sub

Private Sub ListView1_DblClick()
    Call AddToListview
End Sub

'drag drop
Private Sub ListView2_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call AddToListview
End Sub

Private Sub ListView2_DblClick()
    'remove items from listview2
    ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
End Sub


Private Sub mnuAbout_Click()
    MsgBox "Word Counter by Juanito Dado Jr." & vbCrLf & "Please Vote for ME on PSC", , "About"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub CopySelected()
    'Copy the selection and put it on the Clipboard
    Clipboard.Clear
    Clipboard.SetText MSFlexGrid1.Clip
End Sub

'this will fill the background of the selected column with yellow color
'credits goes to gavio of vbforums
Private Sub colBackColor(mfg As MSFlexGrid, col As Long, color As Long)
    With mfg
        .Redraw = False
        .FillStyle = flexFillRepeat
        .col = col
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .CellBackColor = color
        .FillStyle = flexFillSingle
        .col = 1
        .Redraw = True
    End With
End Sub

'background color change
Private Sub Option1_Click(Index As Integer)
         Static lastCol As Long
        If lastCol <> 0 Then
            colBackColor MSFlexGrid1, lastCol, vbWhite
        End If
            lastCol = (Index + 1)
                colBackColor MSFlexGrid1, lastCol, RGB(255, 255, 163)
End Sub
