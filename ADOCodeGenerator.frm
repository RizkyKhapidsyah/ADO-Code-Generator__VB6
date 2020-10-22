VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCodeGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADO - Access 2000 VB Code Generator"
   ClientHeight    =   7500
   ClientLeft      =   1725
   ClientTop       =   930
   ClientWidth     =   9555
   Icon            =   "ADOCodeGenerator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   38
      TabIndex        =   5
      Top             =   7140
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   6960
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
            Picture         =   "ADOCodeGenerator.frx":1042
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ADOCodeGenerator.frx":1156
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ADOCodeGenerator.frx":126A
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ADOCodeGenerator.frx":137E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Select an Access  DB and start the generation"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtCodeGen 
      Height          =   6195
      Left            =   38
      TabIndex        =   3
      Top             =   720
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   10927
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"ADOCodeGenerator.frx":191A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   8242
      TabIndex        =   4
      Top             =   7020
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   7020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblAutor 
      Alignment       =   1  'Right Justify
      Caption         =   "by Carlos Vara"
      Height          =   195
      Left            =   8400
      TabIndex        =   2
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   7140
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmCodeGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-----------------------------------------------------------------+
'| ADO - Access 2000 VB Code Generator                             |
'| copyright ©2001 by Carlos Vara                                  |
'|                                                                 |
'| Author : Carlos A. Vara Pedroza                                 |
'| eMail  : cavarped@ccm.femsa.com.mx; keykorosayork@yahoo.com     |
'|                                                                 |
'| Program Description : I catalogue this program like an utilitie |
'| this utilitie generate a function that you can use inside your  |
'| apps to create an Access 2000 Database With ADO.                |
'|                                                                 |
'| Just copy and paste the code generated to new VB project and    |
'| two references :                                                |
'|                  Microsoft ActiveX Data Objects 2.5 Library     |
'|                  Microsoft ADO Ext. 2.5 for DDL an Security     |
'|                                                                 |
'| If you found any buy, please eMail me!                          |
'| * Don't forget to vote!!! *                                     |
'+-----------------------------------------------------------------+

Option Explicit

Dim cnn    As ADODB.Connection
Dim pth    As String              'Database Path

Dim ErrGen As Boolean

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Private Sub GetPath()
  
  On Error GoTo ErrGetPath
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "All Files(*.*)|*.*|Access 2000 DB (*.mdb)|*.mdb"
  ' Specify default filter
  CommonDialog1.FilterIndex = 2
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  pth = CommonDialog1.FileName
  Exit Sub
  
ErrGetPath:
  'User pressed the Cancel button
  pth = vbNullString
  Exit Sub

End Sub

Private Sub Connect()
  
  On Error GoTo CnnError
  Dim strCnn As String
  
  strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;"
  strCnn = strCnn & "Data Source=" & pth & ";"
  strCnn = strCnn & "Jet OLEDB:Engine Type=5;"
    
  Set cnn = New ADODB.Connection
  cnn.Open strCnn
  Exit Sub

CnnError:
  Dim psw As String

  Select Case Err
    Case Is = -2147217843 'Database password incorrect
      psw = ObtainPassword
      strCnn = vbNullString
      strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;"
      strCnn = strCnn & "Data Source=" & pth & ";"
      strCnn = strCnn & "Jet OLEDB:Engine Type=5;"
      strCnn = strCnn & psw
      If LenB(psw) = 0 Then
        Resume Next
      Else
        Resume
      End If
    Case Else
      MsgBox "Error Number : " & Err & vbCrLf & Error, vbCritical, Err.Source
      End
  End Select

End Sub

Private Sub CodeGen()
  
  On Error GoTo ErrorGen
  Dim Ctl      As ADOX.Catalog
  Dim CtlTbl   As ADOX.Table
  Dim CtlIdx1  As ADOX.Index
  Dim Col      As ADOX.Column
  
  Dim sGen As String
  Dim i    As Integer
  Dim j    As Integer
  Dim k    As Integer
  
  Screen.MousePointer = vbHourglass
    
  txtCodeGen.Text = vbNullString

  'Open the Database Catalog from Actual Connection
  Set Ctl = New ADOX.Catalog
  Ctl.ActiveConnection = cnn
  
  GenerateCode "Private Sub CreateDatabase()"
  GenerateCode "On Error Goto ErrorCreateDB" & vbCrLf & vbCrLf & _
               "Dim Cat     As New ADOX.Catalog" & vbCrLf & _
               "Dim Tbl(" & Ctl.Tables.Count - 1 & ") As ADOX.Table" & vbCrLf & _
               "Dim Idx()   As ADOX.Index" & vbCrLf & _
               "Dim msgErrR As integer" & vbCrLf & _
               "Dim sCnn    As String " & vbCrLf & vbCrLf & _
               "sCnn =  ""Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Jet OLEDB:Engine Type=5;Data Source=" & App.Path & "\NuevaDB.mdb""" & vbCrLf & vbCrLf & _
               "Cat.Create sCnn" & vbCrLf

  ProgressBar1.Max = Ctl.Tables.Count
  ProgressBar1.Visible = True
  lblMessage.Visible = True
  cmdExit.Visible = False
  
  'Get table names
  i = 0
  
  'Table Definitions
  For Each CtlTbl In Ctl.Tables
    
    If CtlTbl.Type = "TABLE" Then
      
      GenerateCode "  '----------* Table Definition of " & CtlTbl.Name & " *----------"
      GenerateCode "  Set Tbl(" & Trim$(Str$(i)) & ")= New ADOX.Table"
      GenerateCode "  Tbl(" & Trim$(Str$(i)) & ").ParentCatalog = Cat"
      GenerateCode "  With Tbl(" & Trim$(Str$(i)) & ")" & vbTab
      GenerateCode "    .Name = """ & CtlTbl.Name & """"
      
      'Field Definitions
      For Each Col In CtlTbl.Columns
        lblMessage.Caption = "Generating code... Table " & CtlTbl.Name
        sGen = sGen & "    .Columns.Append """ & Col.Name & """, " & Trim$(LoadResString(Col.Type)) & IIf(Col.DefinedSize <> 0 And Col.Type <> adBoolean, ", " & Col.DefinedSize, "") & vbCrLf
        
        'Some Properties
        If Col.Properties("AutoIncrement").Value Then
          sGen = sGen & "      .Columns(""" & Col.Name & """).Properties(""AutoIncrement"").Value = True" & vbCrLf
        End If
        If Len(Col.Properties("Description").Value) > 0 Then
          sGen = sGen & "      .Columns(""" & Col.Name & """).Properties(""Description"").Value = """ & Col.Properties("Description").Value & """" & vbCrLf
        End If
        If Not Col.Properties("Nullable").Value = False Then
          sGen = sGen & "      .Columns(""" & Col.Name & """).Properties(""Nullable"").Value = False" & vbCrLf
        End If
        If Len(Col.Properties("Default").Value) > 0 Then
          sGen = sGen & "      .Columns(""" & Col.Name & """).Properties(""Default"").Value = """ & Col.Properties("Default").Value & """" & vbCrLf
        End If
      
      Next Col
      
      sGen = sGen & "  End With"
      
      'Indexes
      If CtlTbl.Indexes.Count > 0 Then
        lblMessage.Caption = "Generating code... Indexes in Table " & CtlTbl.Name
        sGen = sGen & vbCrLf & "  '----------* Index Definitions of " & CtlTbl.Name & " *----------" & vbCrLf
        sGen = sGen & "  ReDim Idx(" & Trim$(Str$(CtlTbl.Indexes.Count - 1)) & ")" & vbCrLf
      End If
      
      j = 0
      
      For Each CtlIdx1 In CtlTbl.Indexes
        lblMessage.Caption = "Generating code... Index " & CtlIdx1.Name & " in Table " & CtlTbl.Name
        sGen = sGen & "  Set Idx(" & Trim$(Str$(j)) & ")= New ADOX.Index" & vbCrLf
        sGen = sGen & "    Idx(" & Trim$(Str$(j)) & ").Name = """ & CtlIdx1.Name & """" & vbCrLf
        If CtlIdx1.PrimaryKey Then
          sGen = sGen & "    Idx(" & Trim$(Str$(j)) & ").PrimaryKey = True" & vbCrLf
        End If
        If CtlIdx1.IndexNulls <> adIndexNullsDisallow Then
          Select Case CtlIdx1.IndexNulls
            Case Is = 0
              sGen = sGen & "    Idx(" & Trim$(Str$(j)) & ").IndexNulls = adIndexNullsAllow" & vbCrLf
            Case Is = 2
              sGen = sGen & "    Idx(" & Trim$(Str$(j)) & ").IndexNulls = adIndexNullsIgnore" & vbCrLf
            Case Is = 4
              sGen = sGen & "    Idx(" & Trim$(Str$(j)) & ").IndexNulls = adIndexNullsIgnoreAny" & vbCrLf
          End Select
        End If
        If CtlIdx1.Unique = True Then
          sGen = sGen & "    Idx(" & Trim$(Str$(j)) & ").Unique = True" & vbCrLf
        End If
        
        If CtlIdx1.Columns.Count = 1 Then
          'Single Column Index
          sGen = sGen & "      Idx(" & Trim$(Str$(j)) & ").Columns.Append """ & CtlIdx1.Columns(0).Name & """" & vbCrLf
          If CtlIdx1.Columns.Item(0).SortOrder = adSortDescending Then
            sGen = sGen & "          Idx(" & Trim$(Str$(j)) & ").Columns(""" & CtlIdx1.Columns(0).Name & """).SortOrder = adSortDescending" & vbCrLf
          End If
        ElseIf CtlIdx1.Columns.Count > 1 Then
          'MultiColumn Index
          For k = 0 To CtlIdx1.Columns.Count - 1
            sGen = sGen & "      Idx(" & Trim$(Str$(j)) & ").Columns.Append """ & CtlIdx1.Columns(k).Name & """" & vbCrLf
            If CtlIdx1.Columns.Item(k).SortOrder = adSortDescending Then
              sGen = sGen & "          Idx(" & Trim$(Str$(j)) & ").Columns(""" & CtlIdx1.Columns(k).Name & """).SortOrder = adSortDescending" & vbCrLf
            End If
          Next k
        End If
        
        j = j + 1
        
      Next CtlIdx1
      
      If j > 1 Then
        sGen = sGen & "  For i = 0 to UBound(Idx)" & vbCrLf
        sGen = sGen & "    Tbl(" & Trim$(Str$(i)) & ").Indexes.Append Idx(i)" & vbCrLf
        sGen = sGen & "  Next i" & vbCrLf
      ElseIf j = 1 Then
        sGen = sGen & "  Tbl(" & Trim$(Str$(i)) & ").Indexes.Append Idx(0)" & vbCrLf
      End If
      
      GenerateCode sGen
      GenerateCode "  Cat.Tables.Append Tbl(" & Trim$(Str$(i)) & ")" & vbCrLf
      
      sGen = vbNullString
      
      
    End If
    i = i + 1
    ProgressBar1.Value = i
    
  Next CtlTbl
  
  'Error code
  sGen = "  Set Cat = Nothing" & vbCrLf
  sGen = sGen & "  Exit Sub" & vbCrLf & vbCrLf
  sGen = sGen & "  ErrorCreateDB:" & vbCrLf
  sGen = sGen & "    msgErrR = MsgBox("""
  sGen = sGen & "    Error No. "" & Err & "" "" & vbCrLf & Error, vbCritical+vbAbortRetryIgnore, ""Code Gen Error"")" & vbCrLf
  sGen = sGen & "    Select Case msgErrR" & vbCrLf
  sGen = sGen & "      Case Is = vbAbort" & vbCrLf
  sGen = sGen & "      If Not (Cat is Nothing) Then" & vbCrLf
  sGen = sGen & "        Set Cat = Nothing" & vbCrLf
  sGen = sGen & "      Endif" & vbCrLf
  sGen = sGen & "      Exit Sub" & vbCrLf
  sGen = sGen & "     Case Is = vbRetry" & vbCrLf
  sGen = sGen & "       Resume Next" & vbCrLf
  sGen = sGen & "     Case Is = vbIgnore" & vbCrLf
  sGen = sGen & "       Resume" & vbCrLf
  sGen = sGen & "    End Select" & vbCrLf & vbCrLf
  sGen = sGen & "End Sub"
  
  GenerateCode sGen

  ProgressBar1.Visible = False
  lblMessage.Visible = False
  cmdExit.Visible = True
  Screen.MousePointer = vbDefault
  Exit Sub

ErrorGen:
  
  MsgBox "Error No. " & Err & vbCrLf & Error, vbCritical, "Error"
  ErrGen = True
  cmdExit.Visible = True
  ProgressBar1.Visible = False
  lblMessage.Visible = False
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
End Sub

Private Sub GenerateCode(ByRef CodeLine As String)
  
  LockWindowUpdate txtCodeGen.hWnd
    txtCodeGen.SelText = CodeLine & vbCrLf
  LockWindowUpdate False
 
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (cnn Is Nothing) Then
    cnn.Close
    Set cnn = Nothing
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  Select Case Button.Index
    Case Is = 1 'Open and Genera
      'Obtain the database path
      GetPath
      'Check if we have correct path!
      If LenB(pth) > 0 And LenB(Dir(pth)) > 0 Then
        'Connect to Database
        Connect
        'Check if we have a connection now
        If Not (cnn Is Nothing) Then
          'Start code generation
          CodeGen
          If Not ErrGen Then
            MsgBox "Generation finished without errors.", vbInformation
          Else
            MsgBox "This program found errors during the generation.", vbExclamation
          End If
        End If
      End If
    Case Is = 3 'Cut
      Clipboard.SetText txtCodeGen.SelText, vbCFText
      txtCodeGen.SelText = vbNullString
    Case Is = 4 'Copy
      Clipboard.SetText txtCodeGen.SelText, vbCFText
    Case Is = 5 'Paste
      txtCodeGen.SelText = Clipboard.GetText(vbCFText)
  End Select
  
End Sub

Private Function ObtainPassword() As String
  
  Dim psw As String
  
  psw = vbNullString
  
  Do While Len(psw) = 0
    frmLogin.Show vbModal
    psw = frmLogin.Password
    If frmLogin.NoMore Then
      ObtainPassword = vbNullString
      Exit Do
    End If
  Loop
  If Len(psw) > 0 And (Not frmLogin.NoMore) Then
    ObtainPassword = ";Jet OLEDB:Database Password=" & psw & ";"
    Unload frmLogin
  Else
    MsgBox "The correct password is needed to" & vbCrLf & "start the code generation.", vbExclamation
    ObtainPassword = vbNullString
    Set cnn = Nothing
    Unload frmLogin
  End If

End Function

