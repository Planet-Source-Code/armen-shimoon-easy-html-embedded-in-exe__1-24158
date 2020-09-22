VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "HTML in file"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "View Internal .html"
      Height          =   555
      Left            =   450
      TabIndex        =   1
      Top             =   4590
      Width           =   2085
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   4155
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   8745
      ExtentX         =   15425
      ExtentY         =   7329
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReturnStr As String
Dim SecondFile As String
Dim FinalStr As String

Private Sub Command1_Click()
ViewWeb 1, 2, web1

End Sub




Private Function ViewWeb(id_name As String, id_second As String, id_browser As WebBrowser)

ReturnStr = StrConv(LoadResString(id_name), vbUnicode)
SecondFile = StrConv(LoadResString(id_second), vbUnicode)

FinalStr = ReturnStr & SecondFile

Open App.Path & "temp.htm" For Output As #1
Print #1, FinalStr
Close #1

web1.Navigate App.Path & "temp.htm"




End Function

Private Sub Form_Unload(Cancel As Integer)
Kill App.Path & "temp.htm"
End Sub
