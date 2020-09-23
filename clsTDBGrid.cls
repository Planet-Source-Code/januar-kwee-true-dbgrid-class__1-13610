VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTDBGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Sequence of using:
'1. SetGridName      name of TdbGrid
'3. SetFormName      just put ME instead of form's name
'4. SetDataControl   optional: set this property if you want to enable sorting

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type TPMPARAMS
   cbSize As Long
   rcExclude As RECT
End Type

Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&

Private Const TPM_RETURNCMD As Long = &H100&
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu&) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpNewItem$) As Long
Private Declare Function ClientToScreen& Lib "user32" (ByVal hwnd&, lpPoint As POINTAPI)
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long

Private WithEvents myGrid As TDBGrid
Attribute myGrid.VB_VarHelpID = -1
Private myDataControl As Data
Private myForm As Form
Dim HasDataControl As Boolean

Property Let SetGridName(TDBGrid1 As TDBGrid)
   Set myGrid = TDBGrid1
   myGrid.AllowUpdate = False
   Set myDataControl = Nothing
   HasDataControl = False
   SaveWithSysHelp = False
End Property

Property Let setFormName(FrmName As Form)
   Set myForm = FrmName
   On Error Resume Next
   'Layout is saved with same name with execute file, except the extension
   myGrid.LayoutFileName = App.Path & "\" & App.EXEName & ".grx"
   myGrid.LayoutName = myForm.Name & myGrid.Name
   myGrid.LoadLayout
   
   'Make a default
   With myGrid
      .Font.Bold = True
      .Font.Name = "Arial"
      .HeadLines = 2
      .HeadForeColor = &H8000000D
      .MarqueeStyle = 2
      .Splits(0).EditBackColor = &HFFFFFF
      .Splits(0).EditForeColor = &HFF0000
      .EvenRowStyle.BackColor = &HC0FFFF
      .OddRowStyle.BackColor = &H80FFFF
      .AlternatingRowStyle = True
   End With
End Property

Property Let setDataControl(mData As Data)
   Set myDataControl = mData
   HasDataControl = True
End Property

Private Sub Class_Initialize()
   HasDataControl = False
   Set myDataControl = Nothing
End Sub

Private Sub myGrid_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
   On Error Resume Next
   SaveLayOut
End Sub

Private Sub SaveLayOut()
   On Error Resume Next
   myGrid.Layouts.Add myForm.Name & myGrid.Name
End Sub

Private Sub myGrid_HeadClick(ByVal ColIndex As Integer)
   SortColumn ColIndex, "ASC"
End Sub

Private Sub SortColumn(ByVal ColIndex As Integer, Order As String)
   Dim mSrc As String
   On Error GoTo Err
   
   'if RecordSource empty, ignore HeadClick
   If Not HasDataControl Then Exit Sub
   mSrc = myDataControl.RecordSource
        
   mSrc = UCase(Trim(mSrc))

   If Left(mSrc, 6) <> "SELECT" Then mSrc = "SELECT * FROM " & mSrc
   If InStr(1, mSrc, "ORDER BY") > 0 Then
      mSrc = Left(mSrc, InStr(1, mSrc, "ORDER BY") - 1)
   End If
   mSrc = mSrc & " ORDER BY [" & myGrid.Columns(ColIndex).DataField & "]"
   myDataControl.RecordSource = mSrc & " " & Order
   myDataControl.Refresh
   
Err:
End Sub

Private Sub myGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = 2 And Shift = 1 Then
      myGrid.Layouts.Remove myForm.Name & myGrid.Name
      myGrid.ClearFields
      myGrid.Refresh
      If HasDataControl Then myDataControl.Refresh
      'MakeStandardTDBGrid myGrid
   ElseIf Button = 2 Then
      Dim hwnd&
      Dim r As RECT
      Dim hPopUp&, hNumber&, hAlignment&, nSelection&, pt As POINTAPI
      hwnd = myGrid.hwnd

      hPopUp = CreatePopupMenu()
      hNumber = CreatePopupMenu()
      hAlignment = CreatePopupMenu()
      'Alignment Menu
      Call AppendMenu(hAlignment, MF_STRING, 11&, ByVal "Left")
      Call AppendMenu(hAlignment, MF_STRING, 12&, ByVal "Right")
      Call AppendMenu(hAlignment, MF_STRING, 13&, ByVal "Center")
      'Number format
      Call AppendMenu(hNumber, MF_STRING, 21&, ByVal "With Parenthesis")
      Call AppendMenu(hNumber, MF_STRING, 22&, ByVal "With Minus Sign")
      Call AppendMenu(hNumber, MF_STRING, 23&, ByVal "Plain format")
      Call AppendMenu(hNumber, MF_STRING, 24&, ByVal "No Decimal")
      
      'Main Menu
      If HasDataControl Then
         Call AppendMenu(hPopUp, MF_STRING, 1&, ByVal "Seek")
         Call AppendMenu(hPopUp, &H800&, 2&, ByVal "-")
         Call AppendMenu(hPopUp, MF_STRING, 31&, ByVal "Sort Ascending")
         Call AppendMenu(hPopUp, MF_STRING, 32&, ByVal "Sort Descending")
         Call AppendMenu(hPopUp, &H800&, 7&, ByVal "-")
      End If
      Call AppendMenu(hPopUp, MF_POPUP, hAlignment, ByVal "Alignment")
      Call AppendMenu(hPopUp, MF_STRING, 3&, ByVal "Format Date")
      Call AppendMenu(hPopUp, MF_STRING, 4&, ByVal "Format String")
      Call AppendMenu(hPopUp, MF_POPUP, hNumber, ByVal "Format Number")
      Call AppendMenu(hPopUp, &H800&, 5&, ByVal "-")
      Call AppendMenu(hPopUp, MF_STRING, 6&, ByVal "Change Caption")
      
      With pt
        .x = x / Screen.TwipsPerPixelX
        .y = y / Screen.TwipsPerPixelY
        Call ClientToScreen(hwnd, pt)
        nSelection = TrackPopupMenu(hPopUp, TPM_RETURNCMD, .x, .y, 1&, hwnd, r)
      End With
      Select Case nSelection
      Case 11, 12, 13
         myGrid.Columns(myGrid.Col).Alignment = nSelection - 11
      Case 21
         myGrid.Columns(myGrid.Col).NumberFormat = "#,##0.00;(#,##0.00);0.00;0.00"
      Case 22
         myGrid.Columns(myGrid.Col).NumberFormat = "#,##0.00 ;#,##0.00-;0.00 ;0.00 "
      Case 23
         myGrid.Columns(myGrid.Col).NumberFormat = "#,##0.00;#,##0.00;0.00;0.00"
      Case 24
         myGrid.Columns(myGrid.Col).NumberFormat = "#,##0;-#,##0;0;0"
      
      Case 3
         myGrid.Columns(myGrid.Col).NumberFormat = "dd mmm yyyy"
      Case 4
         myGrid.Columns(myGrid.Col).NumberFormat = ""
      Case 6
         A = InputBox("Enter Caption", "Metalogic Infomitra", myGrid.Columns(myGrid.Col).Caption)
         If A <> "" Then
            myGrid.Columns(myGrid.Col).Caption = A
         End If
      Case 1
         A = InputBox("Enter " & myGrid.Columns(myGrid.Col).Caption, "Metalogic Infomitra", myGrid.Columns(myGrid.Col).Value)
         If A <> "" Then
            myDataControl.Recordset.FindFirst "[" & myGrid.Columns(myGrid.Col).DataField & "] like '" & A & "*'"
         End If
      
      Case 31
         SortColumn myGrid.Col, "ASC"
      Case 32
         SortColumn myGrid.Col, "DESC"
      End Select
      SaveLayOut
      Call DestroyMenu(hPopUp)
   End If
End Sub

Function Refresh()
   On Error Resume Next
   myGrid.LayoutFileName = App.Path & "\" & App.EXEName & ".grx"
   myGrid.LayoutName = myForm.Name & myGrid.Name
   myGrid.LoadLayout
End Function

