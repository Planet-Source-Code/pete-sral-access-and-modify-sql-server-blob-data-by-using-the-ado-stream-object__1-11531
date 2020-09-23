VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Article ID: Q258038"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   15
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   525
      Left            =   1695
      ScaleHeight     =   465
      ScaleWidth      =   3165
      TabIndex        =   3
      Top             =   420
      Width           =   3225
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   525
      Left            =   1725
      ScaleHeight     =   465
      ScaleWidth      =   3165
      TabIndex        =   2
      Top             =   1140
      Width           =   3225
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save to DB"
      Height          =   495
      Left            =   330
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save to File"
      Height          =   495
      Left            =   315
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Article ID: Q258038
'HOWTO: Access and Modify SQL Server BLOB Data by Using the ADO Stream Object

'--------------------------------------------------------------------------------
'The information in this article applies to:

'ActiveX Data Objects (ADO), version 2.5
'Microsoft Visual Basic Professional and Enterprise Editions for Windows, version 6.0
'Microsoft OLE DB Provider for SQL Server, version 7.0
'Microsoft SQL Server version 7.0

'--------------------------------------------------------------------------------


'SUMMARY
'The Stream object introduced in ActiveX Data Objects (ADO) 2.5 can be used to greatly simplify
'the code that needs to be written to access and modify Binary Large Object (BLOB) data in a
'SQL Server Database. The previous versions of ADO [ 2.0, 2.1, and 2.1 SP2 ] required careful
'usage of the GetChunk and AppendChunk methods of the Field Object to read and write BLOB
'data in fixed-size chunks from and to a BLOB column. An alternative to this method now exists
'with the advent of ADO 2.5. This article includes code samples that demonstrate how the Stream
'object can be used to program the following common tasks:

'Save the data stored in a SQL Server Image column to a file on the hard disk.
'Move the contents of a .gif file to an Image column in a SQL Server table.

Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim mstream As ADODB.Stream

Private Sub Command1_Click()

On Error GoTo Command1_Error

Set cn = New ADODB.Connection
cn.Open "Provider=SQLOLEDB;data Source=PJS-MAIN; Initial Catalog=pubs;User Id=sa;Password="

Set rs = New ADODB.Recordset
rs.Open "Select * from pub_info", cn, adOpenKeyset, adLockOptimistic

Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open
mstream.Write rs.Fields("logo").Value

With cdOpen
    .FileName = ""
    .Filter = "Image (*.gif)|*.gif"
    .ShowSave
    
    If Len(.FileName) <> 0 Then
        mstream.SaveToFile .FileName, adSaveCreateOverWrite
        Picture2.Picture = LoadPicture(.FileName)
    End If
End With
rs.Close
cn.Close

MsgBox "Done saving image from database to " & cdOpen.FileName
Exit Sub
Command1_Error:
    MsgBox Str(Err) & " - " & Error, vbExclamation
    
End Sub

Private Sub Command2_Click()

On Error GoTo Command2_Error

Set cn = New ADODB.Connection
cn.Open "Provider=SQLOLEDB;data Source=PJS-MAIN;Initial Catalog=pubs;User Id=sa;Password="

Set rs = New ADODB.Recordset
rs.Open "Select * from pub_info", cn, adOpenKeyset, adLockOptimistic

Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open

With cdOpen
    .FileName = ""
    .Filter = "Image (*.gif)|*.gif"
    .ShowOpen
    
    If Len(.FileName) <> 0 Then
        
        mstream.LoadFromFile .FileName
        rs.Fields("logo").Value = mstream.Read
        Picture1.Picture = LoadPicture(.FileName)
        rs.Update
    End If
End With

rs.Close
cn.Close
MsgBox "Done saving image " & cdOpen.FileName
Exit Sub
Command2_Error:
    MsgBox Str(Err) & " - " & Error, vbExclamation
End Sub


