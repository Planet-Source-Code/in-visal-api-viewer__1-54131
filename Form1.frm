VERSION 5.00
Begin VB.Form APIV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Viewer (Write by Visal)"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextItem 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4680
      Width           =   8655
   End
   Begin VB.TextBox Text 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   8655
   End
   Begin VB.ListBox VType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox VFunction 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox TypeAPI 
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox ListType 
      Height          =   1260
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox ListConst 
      Height          =   780
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox ListFunction 
      Height          =   2220
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2040
      Width           =   8655
   End
   Begin VB.Label Label4 
      Caption         =   "Selected Items :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Available Items :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Type the first few letters of the word you are looking for:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "API Type :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "APIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ################################
'  Write By Visal
'  Email : visalemail@yahoo.com
'  IRC : irc.uirc.net  in channel #uirc.services and #cambodia
' ################################

' ################################

Dim FunctionChrPos(26) As Integer
' Descript : Store where function that start character A-Z located in list
Dim FunctionCountChr(26) As Integer
' Descript : Count function that start with character A-Z

Dim ConstChrPos(26) As Integer
' Descript : Store where constants that start character A-Z located in list
Dim ConstCountChr(26) As Integer
' Descript : Count constants that start with character A-Z

Dim TypeChrPos(26) As Integer
' Descript : Store where types that start character A-Z located in list
Dim TypeCountChr(26) As Integer
' Descript : Count types that start with character A-Z
Dim InType As Boolean
' Descript : Is in the type or out the type
Dim TypeCode As String
' Descript : Get type

' ################################

Private Sub Form_Load()

InType = False

' Show the waiting form
Form1.Show

' Load API File
LoadData App.Path & "\WIN32API.TXT"

'--------------------------
' Add type of API
TypeAPI.AddItem "Constants"
TypeAPI.AddItem "Declares"
TypeAPI.AddItem "Types"
TypeAPI.Text = "Declares"

End Sub

' ##############################
'  Sub For Loading The API Data
' ##############################
Public Sub LoadData(Filename As String)

Dim strData As String ' Get data in database file
Dim strLine() As String ' Split data in line
Dim LineSplit() As String ' Split line in space
Dim GetConst As String ' Get Const name and variable
Dim GetAsc As Integer ' Get first character asc

'Open File Database
Open Filename For Input As #1
    'Get all database
    strData = StrConv(InputB(LOF(1), 1), vbUnicode)
Close #1

' Split database in line
strLine = Split(strData, vbCrLf)

DoEvents

' Add in list
For i = LBound(strLine) To UBound(strLine)
    
    ' Remark not count in list
    If Left(strLine(i), 1) = "'" Then
        GoTo SkipFor
    End If
    
    ' If this line is in type
    If InType = True Then
        TypeCode = TypeCode & vbCrLf & strLine(i)
    End If
    
    
    If Len(strLine(i)) > 0 Then
        ' Split this line in space
        LineSplit = Split(strLine(i), " ")
        
        '  Check what API type they in
        Select Case LCase(LineSplit(0))
        
            Case "declare"
            
                ' Add it to list
                ListFunction.AddItem LineSplit(2)
                VFunction.AddItem strLine(i)
                
                ' Get first character asc
                GetAsc = Asc(Left(UCase(LineSplit(2)), 1)) - 64
                
                ' Count function that start with first character
                FunctionCountChr(GetAsc) = FunctionCountChr(GetAsc) + 1
                
                ' Reset position
                For j = GetAsc To 26
                    FunctionChrPos(j) = FunctionChrPos(j - 1) + FunctionCountChr(j - 1)
                Next j
                
            Case "const"
                
                ' Get Const name and variable
                GetConst = Right(strLine(i), Len(strLine(i)) - (InStr(strLine(i), " ")))
                
                ' Remove remark out from GetConst
                If InStr(1, GetConst, "'") > 0 Then
                    GetConst = Left(GetConst, InStr(GetConst, "'") - 1)
                End If
                
                ' Add it to list
                ListConst.AddItem GetConst
                
                ' Get first character asc
                GetAsc = Asc(Left(UCase(LineSplit(1)), 1)) - 64
                
                ' Count const that start with first character
                ConstCountChr(GetAsc) = ConstCountChr(GetAsc) + 1
                
                ' Reset position
                For j = GetAsc To 26
                    ConstChrPos(j) = ConstChrPos(j - 1) + ConstCountChr(j - 1)
                Next j
                
            Case "type"
                
                'Add it to list
                ListType.AddItem LineSplit(1)
                
                ' Now we in the type
                InType = True
                
                ' Get Type
                TypeCode = strLine(i)
                
                ' Get first character asc
                GetAsc = Asc(Left(UCase(LineSplit(1)), 1)) - 64
                
                ' Count type that start with first character
                TypeCountChr(GetAsc) = TypeCountChr(GetAsc) + 1
                
                ' Reset position
                For j = GetAsc To 26
                    TypeChrPos(j) = TypeChrPos(j - 1) + TypeCountChr(j - 1)
                Next j
                
            Case "end"
            
                'The type is end
                InType = False
                VType.AddItem TypeCode
                
        End Select
    End If
SkipFor:
Next i

' Hide waiting form
Form1.Hide

End Sub

' #############################

Private Sub ListConst_DblClick()

    ' Add const to Text item
    TextItem.Text = TextItem.Text & vbCrLf & "Const " & ListConst.Text & vbCrLf

End Sub

' #############################

Private Sub ListFunction_DblClick()

    ' Add function to Text item
    TextItem.Text = TextItem.Text & vbCrLf & VFunction.List(ListFunction.SelCount) & vbCrLf

End Sub

' #############################

Private Sub ListType_DblClick()

    ' Add type to Text item
    TextItem.Text = TextItem.Text & vbCrLf & VType.List(ListType.SelCount) & vbCrLf

End Sub

' ############################

Private Sub Text_Change()

' Searching while typing
Select Case LCase(TypeAPI.Text)
    Case "constants"
        Search ListConst, Text.Text, TypeAPI.Text
    Case "declares"
        Search ListFunction, Text.Text, TypeAPI.Text
    Case "types"
        Search ListType, Text.Text, TypeAPI.Text
End Select

End Sub

' #############################

Private Sub TypeAPI_Click()

' Change API type
Select Case LCase(TypeAPI.Text)

    Case "constants"
    
        ' Set Const list position and size
        ListConst.Left = 120
        ListConst.Top = 2040
        ListConst.Visible = True
        ListConst.Height = ListFunction.Height
        ListConst.Width = ListFunction.Width
        
        ListFunction.Visible = False
        ListType.Visible = False
        
        ' Searching
        Search ListConst, Text.Text, TypeAPI.Text
        
    Case "declares"
    
        ' Set function list position and size
        ListFunction.Left = 120
        ListFunction.Top = 2040
        ListFunction.Visible = True
        
        ListType.Visible = False
        ListConst.Visible = False
        
        ' Searching
        Search ListFunction, Text.Text, TypeAPI.Text
        
    Case "types"
    
        ' Set type list position and size
        ListType.Left = 120
        ListType.Top = 2040
        ListType.Visible = True
        ListType.Height = ListFunction.Height
        ListType.Width = ListFunction.Width

        ListFunction.Visible = False
        ListConst.Visible = False
        
        ' Searching
        Search ListType, Text.Text, TypeAPI.Text
        
End Select

End Sub

' ###################################
'  Searching
' ###################################

Private Sub Search(List As ListBox, Word As String, APIType As String)

Dim GetAsc As Integer ' Get first character asc from word you want search
Dim GetPos As Integer ' Get start position of character
Dim GetEnd As Integer ' Get end position of charcter
Dim Found As Boolean ' Found word or not ?

Found = False

' If word is nothing then select first item
If Len(Word) < 1 Then
    List.Selected(0) = True
    Exit Sub
End If

' Get first character asc
GetAsc = Asc(UCase(Left(Word, 1))) - 64

' If in not alphabet then stop search
If GetAsc < 1 Then
    List.Selected(0) = True
    Exit Sub
End If

If GetAsc > 26 Then
    List.Selected(0) = True
    Exit Sub
End If

' Get start and end position by API type
Select Case LCase(APIType)
    Case "constants"
        GetPos = ConstChrPos(GetAsc)
        GetEnd = ConstCountChr(GetAsc) + GetPos
    Case "declares"
        GetPos = FunctionChrPos(GetAsc)
        GetEnd = FunctionCountChr(GetAsc) + GetPos
    Case "types"
        GetPos = TypeChrPos(GetAsc)
        GetEnd = TypeCountChr(GetAsc) + GetPos
End Select

' If there are no item for it then stop search
If GetPos = GetEnd Then
    List.Selected(0) = True
    Exit Sub
End If

'Searching
For i = GetPos To GetEnd
    If InStr(1, UCase(List.List(i)), UCase(Word)) = 1 Then
        List.Selected(i) = True
        Found = True
        Exit For
    End If
Next i

End Sub
