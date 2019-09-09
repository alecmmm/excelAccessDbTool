VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccessDBUserForm 
   Caption         =   "CMRM Database Tool"
   ClientHeight    =   4880
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6180
   OleObjectBlob   =   "AccessDBUserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "AccessDBUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Developed by Alec McKay. December 2017

Private list() As Variant
Private item As Variant
Private database As databaseHolder
Dim rangeSelection As String

Private Sub AccessExcelButton_Click()

database.copyTableInto AccessExcelBox.value

End Sub

Private Sub AccessExcelButton2_Click()

database.copyTable AccessExcelBox2.value

End Sub

Private Sub deleteCommandButton_Click()

database.deleteTable deleteTableBox.value

End Sub

Private Sub headerCheckbox1_Click()

database.setCheckBox1 = headerCheckBox1.value

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()

Set database = New databaseHolder

database.Init

list = database.getList

For Each item In list

If InStr(item, "MSys") = 0 And InStr(item, "~TMP") = 0 Then

deleteTableBox.AddItem item
AccessExcelBox.AddItem item
AccessExcelBox2.AddItem item

End If

FileTextBox1.Text = "Database: " & database.getDBName

FileTextBox2.Text = "Database: " & database.getDBName


Next

FileTextBox3.Text = "Database: " & database.getDBName

End Sub
f��u              ! � P�H��              ! �    "  ��u          " 
 ��p�u              "  ��P    8��  "  �3P   #  �W�          #  p��W�              # 3 �x��              # E 
  
� �Mg�'$  � ����A        �