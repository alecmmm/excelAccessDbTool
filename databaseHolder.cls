VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "databaseHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Developed by Alec McKay. December 2017

' Class name: databaseHolder

'Important: Excel pulls columns in the order they are ordered in design view. _
When working with a new table, make sure that the order of columns in datasheet _
view and design view is the same.


Public dataBase1 As String
Dim appAccess As New Access.Application
Dim list As Variant
Dim lst As Variant
Dim dbPath As String
Dim strQry As String
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim headerCheckBox1 As Boolean

Public Sub Init()

'if no Access database is currently open, go to error handler to open Access database

On Error GoTo NoAccessFile
    
    'find access database that's currently open and set
    
    dataBase1 = CurrentProject.Path & "\" & CurrentProject.Name
    Set appAccess = Access.Application
    
    'make database visible
    
    If appAccess.Visible = False Then
        appAccess.OpenCurrentDatabase dataBase1
        appAccess.Visible = True
    End If
    list = tableNames
    
    'if no Access database is fount, prompt user to open one
    
NoAccessFile:
    If Err.number = 2467 Then
        
        MsgBox "Please open Access database."
        
        dataBase1 = Application.GetOpenFilename _
        (Title:="Please choose an Access file to open", _
        FileFilter:="Excel Files *.accdb* (*.accdb*),")
        

    End If
    
    'set "Data contains headers?" checkbox to true by default
    
    headerCheckBox1 = True
    
    headerCheckBox2 = True
    
End Sub

'setter for "Data contains headers?" checkbox
Property Let setCheckBox1(header As Boolean)

headerCheckBox1 = header

End Property

'getter for Access database name
Property Get getDBName() As String

getDBName = CurrentProject.Name

End Property

'getter for list of tables
Property Get getList() As Variant

getList = list

End Property

'Sub that copies Access table into Excel

Public Sub copyTable(strQry As String)
  
  If strQry = "" Then
  
  MsgBox "Choose a table to export before clicking."
  
  Exit Sub
  
  End If
  
  MsgBox "Please note that columns will be in the same order as the fields in Design View."
  
  'create ADODB connection to connect Access to Excel
  
ActiveSheet.Select
    Rows("2:1048576").Select 'Select row 2 to row 1048576
    Selection.ClearContents 'Delete content from row 2 to row 1048576 - should be good for a few years (now it's 2014)
    Range("A1").Select
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataBase1 & ";"
    Set rs = New ADODB.Recordset
    
    With rs
      Set .activeconnection = cn
      .Open "SELECT * FROM " & strQry
      
      Range("A1").Select
      
      For i = 0 To .Fields.Count - 1
    ActiveCell.Offset(0, i) = .Fields(i).Name
    Next i
    
    End With
    Range("A2").CopyFromRecordset rs
    rs.Close
    cn.Close
    
    MsgBox "Done importing " & strQry & " into Excel worksheet"

End Sub

'Sub that copies Excel tables into Access

Public Sub copyTableInto(strQry As String)
    
    On Error GoTo ErrorHandler
    
    'Close table
    DoCmd.Close acTable, strQry, acSaveYes
    
'If headers are not included in data, then copy headers from Access and paste to Excel
    
If headerCheckBox1 = False Then insertHeaders strQry
    
 'create Access connection to connect Excel to Access
    
  Set cn = CreateObject("ADODB.Connection")
  dbWb = Application.ActiveWorkbook.FullName
  dbWs = Application.ActiveSheet.Name
  scn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataBase1
  dsh = "[" & Application.ActiveSheet.Name & "$]"
  cn.Open scn

  sSQL = "INSERT INTO " & strQry & "([])"
  sSQL = sSQL & "SELECT * FROM [Excel 8.0;HDR=" & headerCheckBox1 & ";DATABASE=" & dbWb & "]." & dsh

  cn.Execute sSQL
  
  cn.Close

    
    MsgBox "Done importing Excel data into " & strQry
    
    Exit Sub
    
    'if Excel file hasn't been saved, then conneciton can't be established
    
ErrorHandler:
    If Err.number = -2147217865 Then
    '3011
        MsgBox "Save Excel file before clicking."
    Exit Sub
    
    Else
        MsgBox Err.Description
    
    End If
    
    Exit Sub

End Sub

'sub to clear data in Access table

Public Sub deleteTable(strQry As String)
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ConnectionString As String
Dim strQry2 As String
    
    If strQry = "" Then
    
        MsgBox "Choose an Access table to clear before clicking."
        Exit Sub
        
    End If
    
    'Close table
    DoCmd.Close acTable, strQry, acSaveYes

'Setup the connection string for accessing MS SQL database
    ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataBase1 & ";"

    'Opens connection to the database
    cn.Open ConnectionString
    'Timeout error in seconds for executing the entire query; this will run for 15 minutes before VBA timesout, but your database might timeout before this value
    cn.CommandTimeout = 30

    'This is your actual MS SQL query that you need to run; you should check this query first using a more robust SQL editor (such as HeidiSQL) to ensure your query is valid
    strQry2 = "DELETE * FROM " & "[" & strQry & "]"

    'Performs the actual query
    rs.Open strQry2, cn
    'Dumps all the results from the StrQuery into cell A2 of the first sheet in the active workbook
    'Sheets(1).range("A2").CopyFromRecordset rst
    MsgBox strQry & " has been cleared."
    
End Sub

'Sub inserts headers from Access into Excel

Private Sub insertHeaders(strQry As String)
        
        MsgBox "Please note that headers will be in the same order as the fields in Design View."
        
        Range("A1").EntireRow.Insert
        
        'setup connection
        
        Set cn = New ADODB.Connection
        cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataBase1 & ";"
        Set rs = New ADODB.Recordset
        With rs
            Set .activeconnection = cn
            .Open "SELECT * FROM " & "[" & strQry & "] WHERE 1 = 0" 'temporary fix as the code seems to need a WHERE clause to use the Open method
            For i = 0 To .Fields.Count - 1
            
            Range("A1").Select
            ActiveCell.Offset(0, i) = .Fields(i).Name
            Next i
         
        End With
        rs.Close
        cn.Close
        
End Sub

'function to populate the table names from the Access database

Function tableNames() As Variant

    Dim list() As Variant
    Dim counter As Integer
    
    Dim db As DAO.database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Set db = CurrentDb
    
    counter = 0
    For Each tdf In db.TableDefs
        ' ignore system and temporary tables
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
            Debug.Print tdf.Name
        End If
        
        ReDim Preserve list(counter)
        list(counter) = tdf.Name
        counter = counter + 1
        
    Next
    
        'counter = counter + 1
    
    For Each qdf In db.QueryDefs
        If Not (qdf.Name Like "MSys*" Or qdf.Name Like "~*") Then
            Debug.Print qdf.Name
        End If
        
        ReDim Preserve list(counter)
        list(counter) = qdf.Name
        counter = counter + 1
        
    Next
    
    Set tdf = Nothing
    Set qdf = Nothing
    Set db = Nothing
    
    tableNames = list

End Function


                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                