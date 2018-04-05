' ===================================================
'
' Author	: Christophe Avonture
' Date		: March 2018
'
' Connect to a SQL Server database, obtain the list of
' tables in that db (process all schemas), get the structure
' of each tables thanks the sp_columns stored procedure and
' for each table, export that structure in a results subfolder
'
' At the end, we'll have as many files as there are tables in
' the database. One .CSV file by table.
'
' The content of the CSV will be what is returned by the sp_columns
' stored procedure.
'
' Documentation : https://github.com/cavo789/sql_sp_colums
' ===================================================

Option Explicit

Const cServerName = "" 		' <== Name of your SQL server
Const cDatabaseName = ""	' <== Name of the database
Const cUserName = ""		' <== User name
Const cPassword = ""		' <== User password

Dim sDatabaseName, sServerName, sUserName, sPassword

' ---------------------------------------------------
'
' Show help screen
'
' ---------------------------------------------------
Sub ShowHelp()

	wScript.echo " ======================================="
	wScript.echo " = sp_columns for SQL Server databases ="
	wScript.echo " ======================================="
	wScript.echo ""
	wScript.echo " This script requires four parameters : the server, "
	wScript.echo " database name, login and password to use for the connection."
	wScript.echo ""
	wScript.echo " " & wScript.ScriptName & " 'ServerName', 'dbTest', 'Login', 'Password'"
	wScript.echo ""
	wScript.echo "To get more info, please read https://github.com/cavo789/sql_sp_colums"
	wScript.echo ""

	wScript.quit

End sub

' ---------------------------------------------------
'
' Return the current, running, folder
'
' ---------------------------------------------------
Public Function getCurrentFolder()

	Dim objFSO, objFile
	Dim sFolder

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(Wscript.ScriptFullName)

	sFolder = objFSO.GetParentFolderName(objFile) & "\"

	Set objFile = Nothing
	Set objFSO = Nothing

	getCurrentFolder = sFolder

End Function

' ---------------------------------------------------
'
' Create a folder if not yet there
'
' ---------------------------------------------------
Public Function makeFolder(ByVal sFolderName)

	Dim objFSO

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If Not objFSO.FolderExists(sFolderName) Then
		Call objFSO.CreateFolder(sFolderName)
	End if

	Set objFSO = Nothing

End Function

' ---------------------------------------------------
'
' Remove all files in the specified folder
'
' ---------------------------------------------------
Public Function emptyFolder(ByVal sFolderName)

	Dim objFSO, objFiles, objFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set objFiles = objFSO.GetFolder(sFolderName).Files

	For Each objFile In objFiles
		objFile.Delete
	Next

	set objFile = Nothing
	set objFSO = Nothing

End function

' ---------------------------------------------------
'
' Create a text file on the disk, UTF-8 with LF
'
' ---------------------------------------------------
Public Sub CreateTextFile(ByVal sFileName, ByVal sContent)

	Dim objStream

	Set objStream = CreateObject("ADODB.Stream")

	With objStream
		.Open
		.CharSet = "x-ansi" ' "UTF-8"
		.LineSeparator = 10
		.Type = 2 ' adTypeText
		.WriteText sContent
		.SaveToFile sFileName, 2
		.Close
	End with

	set objStream = Nothing

End Sub

Dim sDSN, sSQL
Dim objConn, rsTables, rs, fld
Dim sLine, sPath, sFileName, sTableName, sCSV, sContent, sMDTable
Dim wFilesCount

	' Get constants
	sServerName = trim(cServerName)
	sDatabaseName = trim(cDatabaseName)
	sUserName = trim(cUserName)
	sPassword = trim(cPassword)

	' If one variable is not set by constants, get from
	' command line arguments
	If (sServerName = "") or (sDatabaseName = "") or _
		(sUserName = "") or (sPassword = "") Then
		If (wScript.Arguments.Count < 4) Then
			Call ShowHelp
			wScript.quit
		Else
			' Read parameters server -> db -> login -> password
			sServerName = Trim(Wscript.Arguments.Item(0))
			sDatabaseName = Trim(Wscript.Arguments.Item(1))
			sUserName = Trim(Wscript.Arguments.Item(2))
			sPassword = Trim(Wscript.Arguments.Item(3))
		End if
	End If

	' Define the results folder : a subfolder of the folder
	' containing this VBS script.
 	sPath = getCurrentFolder() & "results\"
	makeFolder(sPath)

	' Remove files from a previous run
	emptyFolder(sPath)

	wFilesCount = 0

	' Define the connection string
	sDSN = "Driver={SQL Server};Server={" & sServerName & "};" & _
		"Database={" & sDatabaseName & "};" & _
		"User Id={" & sUserName & "};" & _
		"Password={" & sPassword & "};"

	Set objConn = CreateObject("ADODB.Connection")
	Set rsTables = CreateObject("ADODB.Recordset")

	objConn.ConnectionTimeout = 60
	objConn.CommandTimeout = 60

	objConn.Open sDSN

	' Get the list of tables in the database
	sSQL = "SELECT TABLE_SCHEMA As [Schema], " & _
		"TABLE_NAME As [TableName] " & _
		"FROM " & sDatabaseName & ".INFORMATION_SCHEMA.TABLES " & _
		"WHERE TABLE_TYPE = 'BASE TABLE' " & _
		"ORDER BY TABLE_SCHEMA, TABLE_NAME;"

	Set rsTables = objConn.Execute(sSQL)

	If Not rsTables.Eof Then
		' Iterate for each table

		Set rs = CreateObject("ADODB.Recordset")

		Do While Not rsTables.EOF

			' Retrieve the list of columns in each table

			sSQL = "sp_columns " & _
				rsTables.Fields("TableName").Value & ", " & _
				rsTables.Fields("Schema").Value & ", " & _
				sDatabaseName

			Set rs = objConn.Execute(sSQL)

			If Not rs.Eof Then

				sCSV = ""
				sMDTable = ""
				sLine = ""

				' Derive the filename :
				'	* The database name (f.i. dbAdmin)
				'	* The schema (f.i. dbo)
				'	* The table name (f.i. tblName)
				sTableName = rs.Fields("TABLE_QUALIFIER").Value & _
					"." & rs.Fields("TABLE_OWNER").Value & _
					"." & rs.Fields("TABLE_NAME").Value

				sFileName = sPath & replace(sTableName, ".", "_") & ".csv"

				' Get the list of headers, the list of columns
				For Each fld In rs.Fields
					sLine = sLine & fld.Name & ";"
				Next

				sLine = left(sLine, Len(sLine) -1)

				' Prepare the CSV content with, as first line,
				' the column's headers
				sCSV = sLine & vbCrLf

				' And do the same for the Markdown table
				' Prepare the table declaration

				sMDTable = "| # | Name | Type | Length | " & _
				 	"IsNullable |" & vbLF & _
					"| --- | --- | --- | --- | --- |" & vbLF

				Do While Not rs.EOF

					sLine = ""
					For Each fld In rs.Fields
						sLine = sLine & fld.Value & ";"
					Next

					sLine = left(sLine, Len(sLine) -1)

					' sLine is a data row, add it into the CSV content
					sCSV = sCSV & sLine & vbCrLf

					sMDTable = sMDTable & _
						"| " & rs.Fields("Ordinal_Position").Value & _
						" | " & rs.Fields("Column_Name").Value & _
						" | " & rs.Fields("Type_Name").Value & _
						" | " & rs.Fields("Precision").Value & _
						" | " & rs.Fields("Is_Nullable").Value & _
						" |" & vbLf

					rs.MoveNext

				Loop

			End if

			sCSV = left(sCSV, Len(sCSV) - len(vbCrLf))

			rs.Close
			Set rs = Nothing

			' Create the file
			Call CreateTextFile(sFileName, sCSV)
			wFilesCount = wFilesCount + 1

			' --------------------------------------
			' For documentation only purposes
			' Create a .md file by table
			' Draw a table within the .md file with very columns
			' and add an hyperlink to the .csv file to get
			' the full table
			sFileName = sPath & replace(sTableName, ".", "_") & ".md"
			sContent = "# Table structure" & vbLf & vbLf & _
				"%LASTUPDATE%" & vbLf & vbLf & _
				sMDTable & vbLf & _
				"[Full description](%URL%.files/" & replace(sTableName, ".", "_") & ".csv)" & vbLf
			Call CreateTextFile(sFileName, sContent)
			' --------------------------------------

			rsTables.MoveNext

		Loop

	End if

	rsTables.Close

	Set rsTables = Nothing
	Set objConn = Nothing

	If (wFilesCount > 0) Then
		wScript.echo wFilesCount & " files have been " &_
			"created in " & sPath
	End If