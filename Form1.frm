VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Report Descriptor Import tool"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtCATCFile 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton btnPickCATCFile 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton btnImportFROMCATC 
      Caption         =   "Import from CATC"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton btnLaunchDT 
      Caption         =   "Launch Descriptor Tool"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton PickFile1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox FileName1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ConvertToHIDFile"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ST7map.map, false, RAM_DEF.H


'M:\Eserve\FW\2511 OmniSmart G8.0\code\x056\x056a\proj.map, false, M:\Eserve\FW\2511 OmniSmart G8.0\code\x056\x056a\c_ramdef.h
'M:\Eserve\FW\2553 150LMR MicroPower LiIon\code\x006\x006v\proj.map, false, M:\Eserve\FW\2553 150LMR MicroPower LiIon\code\x006\x006v\c_ramdef.h


#If Win32 Then
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

#Else
Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, _
ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, _
ByVal lpszDir$, ByVal fsShowCmd%) As Integer

Declare Function GetDesktopWindow Lib "USER" () As Integer
#End If

Dim OutFileName As String
Dim AppendSymbols As Boolean
Dim CHeaderParsed As Boolean

Private Type Symbol
    Symbol As String
    DataType As String
    Comment As String
End Type

Private Sub btnImportFROMCATC_Click()

    'On Error GoTo ReadErrorHandle

    'FileSystemObject must reference Microsoft Scripting Engine:
    Dim fs As New FileSystemObject
    Dim InFile As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set InFile = fs.OpenTextFile(txtCATCFile.Text, 1, False, 0)
    Dim FileLine As String

    Dim TextItems(2000) As String
    'Dim DataItemIndex As Long
    Dim DataItems(2000) As String      'Supports at most 2000 items.
    
    EndOfFile = False
    DataItemIndex = 0
    linenumber = 0
    
    Do While InFile.AtEndOfStream <> True
        FileLine = InFile.ReadLine

        label = InStrRev(FileLine, "------------------------------------------------------------------------------------------------------------")
        If label <> 0 Then
            GoTo STARTGETTINGDATA
        End If

    Loop
    
    
    Dim Bytes As String
    Dim Bytes2 As String
    Dim Text As String
    
STARTGETTINGDATA:
    While InFile.AtEndOfStream <> True
        FileLine = InFile.ReadLine
        
        label = InStrRev(FileLine, vbTab)
        If label <> 0 Then
            Bytes = Mid(FileLine, label, Len(FileLine) - label)
        Else
            'No TAB!!
            'MsgBox "No tab" & FileLine
            label = InStrRev(FileLine, ")")
            Bytes = Mid(FileLine, label + 1, Len(FileLine) - label - 1)
        End If
        Bytes = StripWhiteSpaceFromEndOfString(Bytes)
        Bytes = StripWhiteSpaceFromStartOfString(Bytes)
        
        Text = Left(FileLine, label)
        Text = StripWhiteSpaceFromEndOfString(Text)
        'Text = StripWhiteSpaceFromStartOfString(Text)      'Keep the tabbing to keep track of collections.
        
        Dim Parenthesis As Integer
        Parenthesis = InStr(FileLine, "(")
        If (Parenthesis <> 0) Then
            Parenthesis = InStr(FileLine, ")")
            If (Parenthesis = 0) Then
                'GARBAGE OVERFLOW LINE.  IGNORE THIS:
                FileLine = InFile.ReadLine
                Parenthesis = InStr(FileLine, ")")
                If (Parenthesis <> 0) Then
'                    Bytes2 = Left(FileLine, Parenthesis - 1)
'                    Bytes2 = StripWhiteSpaceFromEndOfString(Bytes2)
'                    Bytes2 = StripWhiteSpaceFromStartOfString(Bytes2)
'                    Bytes = Bytes & " " & Bytes2
                Else
                    MsgBox "ERROR! Unmatched parenthesis"
                End If
            End If
        End If
        
        If label <> 0 Then
            'Line contains data
            
            DataItems(DataItemIndex) = Bytes
            TextItems(DataItemIndex) = Text
            
            DataItemIndex = DataItemIndex + 1
        End If
    Wend


    Dim OutFileName As String
    
    With CommonDialog1
        .DialogTitle = "Save as .h File"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .InitDir = VB.App.Path
        .Flags = cdlOFNHideReadOnly
        .Filter = "h Files (*.h)|*.h|All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
    '    if .Action
        
        OutFileName = .FileName
        .FileName = ""
    End With
    
    Set OutFile = fs.CreateTextFile(OutFileName, True)
    
    OutFile.WriteLine "rom struct{byte report[HID_RPT_UPS_SIZE];}hid_rpt_UPS={"
    
    Dim OutputLength As Integer
    Dim OutLine As String
    Dim i As Integer
    For i = 0 To DataItemIndex - 1
        OutLine = ""
        OutLine = OutLine & "0x" & Replace(DataItems(i), " ", ", 0x") & "," ' & vbTab & vbTab & "//" & TextItems(i)
        OutputLength = Len(OutLine)
        If (OutputLength < 56) Then
            If (OutputLength Mod 4 <> 0) Then
                OutLine = OutLine & vbTab
                OutputLength = Int(OutputLength / 4) * 4 + 4
            End If
            
            'The length is now evenly divisble by 4.  Add the correct number of tabs to get to the desired line length
            Dim j As Integer
            For j = 1 To (56 - OutputLength) / 4
                OutLine = OutLine & vbTab
            Next
        End If
        
        OutLine = OutLine & "//" & TextItems(i)
        
        OutFile.WriteLine OutLine
    Next
    OutFile.WriteLine "};"
    
    OutFile.Close
    
    
ReadErrorHandle:
    

End Sub

Private Sub btnLaunchDT_Click()
    SHELL """" & VB.App.Path & "\dt.exe" & """", vbNormalFocus
End Sub

Private Sub btnPickCATCFile_Click()
    With CommonDialog1
        .DialogTitle = "Open txt File"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Flags = cdlOFNHideReadOnly
        .Filter = "Source Files (*.c,*.h, *.txt)|*.c;*.h;*.txt|Source Files (*.c)|*.c|(*.txt)|*.txt|(*.h)|*.h|All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
    '    if .Action
        
        txtCATCFile.Text = .FileName
        .FileName = ""
    End With
End Sub

Private Sub Form_Load()
    Dim a_strArgs() As String
    Dim CmdLineOpts As String
    
    
    OutFileName = VB.App.Path & "\output.hid"
    
    Me.Show
        
End Sub


Private Sub Command1_Click()
    Dim i As Integer

    Dim fs As New FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If (Not fs.FileExists(FileName1.Text)) Then Exit Sub
    
    ImportDescriptorData FileName1.Text
    
End Sub


Private Function ImportDescriptorData(InFileName As String) As Integer
   
    On Error GoTo ReadErrorHandle

    'FileSystemObject must reference Microsoft Scripting Engine:
    Dim fs As New FileSystemObject
    Dim InFile As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set InFile = fs.OpenTextFile(InFileName, 1, False, 0)
    Dim FileLine As String
    Dim BlockComment As Boolean

    Dim DataItems() As String
    'Dim DataItemIndex As Long
    Dim DTData(2000) As String      'Supports at most 2000 items.
    
    EndOfFile = False
    DataItemIndex = 0
    linenumber = 0
    BlockComment = False
    
    While InFile.AtEndOfStream <> True
        FileLine = InFile.ReadLine

        If (BlockComment) Then
        
            label = InStr(FileLine, "*/")
            If label <> 0 Then
                FileLine = Right(FileLine, Len(FileLine) - label - 1)    'strip comment lines
                BlockComment = False
            Else
                FileLine = ""
            End If
        End If
        
        If (Not BlockComment) Then
            label = InStr(FileLine, "//")
            If label <> 0 Then
                FileLine = Left(FileLine, label - 1)      'strip comment lines
            End If
            
            label = InStr(FileLine, "/*")
            If label <> 0 Then
            
                label2 = InStr(FileLine, "*/")
                If (label2 <> 0) Then
                    FileLine = Left(FileLine, label - 1) & Mid(FileLine, label2 + 2, Len(FileLine) - label2 - 1) 'strip comment lines
                Else
                    BlockComment = True
                    FileLine = Left(FileLine, label - 1)      'strip the rest of the line
                End If
            End If
        End If
        
        FileLine = StripWhiteSpaceFromEndOfString(FileLine)
        FileLine = StripWhiteSpaceFromStartOfString(FileLine)
        
        label = InStr(FileLine, ",")
        If label <> 0 Then
            'Line contains data
            
            Dim NBytesOnLine As Integer
            
            NBytesOnLine = 0
            DataItems = Split(FileLine, ",")
            Dim i As Integer
            For i = 0 To UBound(DataItems) - 1
            
                DataItems(i) = StripWhiteSpaceFromEndOfString(DataItems(i))
                DataItems(i) = StripWhiteSpaceFromStartOfString(DataItems(i))
                If (DataItems(i) <> "") Then
                    If (i = 0) Then DTData(DataItemIndex) = ""
                    NBytesOnLine = NBytesOnLine + 1
                    DTData(DataItemIndex) = DTData(DataItemIndex) & ValueToHex(DataItems(i))        'Convert this to HEX 2 character width
                End If
            Next
            For i = NBytesOnLine * 2 + 1 To 19 Step 2
                DTData(DataItemIndex) = DTData(DataItemIndex) & "00"
            Next i
            DataItemIndex = DataItemIndex + 1
        End If
    Wend
    
    Dim NItems_L As Integer
    Dim NItems_H As Integer
    
    NItems_L = H2D(FXD(D2H(DataItemIndex), 2))
    NItems_H = H2D(Left(FXD(D2H(DataItemIndex), 4), 2))
    
    On Error GoTo WriteErrorHandle
    
    Dim nFileNum As Integer
        
    With CommonDialog1
        .DialogTitle = "Save as .HID File"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .InitDir = VB.App.Path
        .Flags = cdlOFNHideReadOnly
        .Filter = "HID Files (*.hid)|*.hid|All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Function
        End If
    '    if .Action
        
        OutFileName = .FileName
        .FileName = ""
    End With
    
    ' Delete contents of Filename because the entire file is in memory
    nFileNum = FreeFile
    Open OutFileName For Output As #nFileNum
    Close #nFileNum ' Close the file before opening in Binary mode
    
    ' Open the file
    nFileNum = FreeFile
    Open OutFileName For Binary Lock Read Write As #nFileNum
    
    ' Save header information:
    Put #nFileNum, , ToByte(H2D("22"))
    Put #nFileNum, , ToByte(H2D("00"))
    Put #nFileNum, , ToByte(H2D("01"))
    Put #nFileNum, , ToByte(H2D("00"))
    Put #nFileNum, , ToByte(H2D("FE"))
    Put #nFileNum, , ToByte(H2D("CA"))
    Put #nFileNum, , ToByte(NItems_L)
    Put #nFileNum, , ToByte(NItems_H)
    Put #nFileNum, , ToByte(H2D("0A"))
    For i = 1 To 25
        Put #nFileNum, , ToByte(H2D("00"))
    Next i
    
    Dim UsagePage As String
    UsagePage = "00"
    Dim NewDTData As String
    
    'Save Data:
    For i = 0 To DataItemIndex - 1
        If (Left(DTData(i), 2) = "05") Then
            UsagePage = Mid(DTData(i), 3, 2)
        End If
        If (Left(DTData(i), 2) = "09") Then
            NewDTData = Left(DTData(i), 6)
            NewDTData = NewDTData & UsagePage & Mid(DTData(i), 8, Len(DTData(i)) - 8)
            DTData(i) = NewDTData
        End If
        For j = 1 To 19 Step 2
            Put #nFileNum, , ToByte(H2D(Mid(DTData(i), j, 2)))
        Next j
    Next i
    
    Close #nFileNum
    
    Exit Function

ReadErrorHandle:
    MsgBox "Error parsing file:  " & Err.Description & vbCrLf & FileLine
    Exit Function
WriteErrorHandle:
    MsgBox "Error Writing file:  " & Err.Description
    Exit Function
End Function

Private Function ToByte(inputval As Variant) As Byte
    ToByte = inputval
End Function


Private Function ValueToHex(inputval As String)
    Dim label As Integer
    
    On Error GoTo ConversionError
    
    label = InStr(inputval, "0x")
    If label <> 0 Then
        ValueToHex = FXD(Mid(inputval, 3, Len(inputval) - 2), 2)
    Else
        ValueToHex = FXD(D2H(inputval), 2)
    End If
    
    Exit Function
    
ConversionError:
    ValueToHex = 255
End Function

Public Function D2H(ByVal DecVal As Long) As String
    D2H = Hex$(DecVal)
End Function

'This converts hex values to decimal.  (Note that it is a signed result)
Public Function H2D(ByVal HexVal As Variant) As Variant
    If (Len(HexVal) > 8) Then
        H2D = -1        'Prevent Overflow error.
        Exit Function
    End If
    H2D = Val("&H" + HexVal)
End Function

'This function is intended to format a hex value to a specified
'number of characters by adding leading 0's or truncating leading characters.
Public Function FXD(LongVar As Variant, Width As Variant) As Variant
    Dim TempWork As String
    Dim TempWidth As Long
    
    TempWidth = CLng(Width)
    TempWork = LongVar
    
    If Len(TempWork) > TempWidth Then
        'FXD = Mid(TempWork, Len(TempWork) - TempWidth, TempWidth)
        FXD = Right(TempWork, TempWidth)
        Exit Function
    End If
    
    Do Until Len(TempWork) = TempWidth
        TempWork = "0" & TempWork
    Loop
    
    FXD = TempWork
End Function

Private Function StripWhiteSpaceFromStartOfString(instring As String)
    Dim workingstring As String
    
    workingstring = instring
    For i = 1 To Len(workingstring) Step 1
        If (Mid(workingstring, i, 1) = " " Or _
            Mid(workingstring, i, 1) = vbTab Or _
            Mid(workingstring, i, 1) = ";") Then
        Else
            'Found non-whitespace
            Exit For
        End If
    Next
    workingstring = Right(workingstring, Len(instring) - i + 1)
    StripWhiteSpaceFromStartOfString = workingstring
End Function

Private Function StripWhiteSpaceFromEndOfString(instring As String)
    Dim workingstring As String
    
    workingstring = instring
    For i = Len(workingstring) To 1 Step -1
        If (Mid(workingstring, i, 1) = " " Or _
            Mid(workingstring, i, 1) = vbTab Or _
            Mid(workingstring, i, 1) = ";") Then
        Else
            'Found non-whitespace
            Exit For
        End If
    Next
    workingstring = Left(workingstring, i)
    StripWhiteSpaceFromEndOfString = workingstring
End Function

Private Sub PickFile1_Click()
    With CommonDialog1
        .DialogTitle = "Open txt File"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Flags = cdlOFNHideReadOnly
        .Filter = "Source Files (*.c,*.h, *.txt)|*.c;*.h;*.txt|Source Files (*.c)|*.c|(*.txt)|*.txt|(*.h)|*.h|All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
    '    if .Action
        
        FileName1.Text = .FileName
        .FileName = ""
    End With
           
End Sub


'This function strips leading spaces from any string
Public Function StripLeadingSpaces(str As String)
    Dim i As Integer
    On Error Resume Next
    'strip leading spaces:
    For i = 1 To Len(str)
        If Mid(str, 1, 1) = " " Or Mid(str, 1, 1) = vbTab Then
            str = Right(str, Len(str) - 1)
        Else
            Exit For
        End If
    Next
    StripLeadingSpaces = str
End Function

'This function strips trailing spaces from any string
Public Function StripTrailingSpaces(str As String)
    Dim i As Integer
    On Error Resume Next
    'strip leading spaces:
    For i = 1 To Len(str)
        If Right(str, 1) = " " Or Right(str, 1) = vbTab Then
            str = Left(str, Len(str) - 1)
        Else
            Exit For
        End If
    Next
    StripTrailingSpaces = str
End Function


'This function can be used to open any file for which there is an association (excel, word, etc), or any executable file
Public Function StartDoc(DocName As String) As Long
 Dim Scr_hDC As Long
 Scr_hDC = GetDesktopWindow()
 StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function

'This function can be used to execute any command line operation
Public Sub ShellExec(cmdline As String)
    SHELL cmdline, vbNormalFocus
End Sub
