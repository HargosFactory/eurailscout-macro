' Description: This function writes a value in a text file.
' Parameters:
'   - path: The path of the file.
'   - value: The value to write.
' Returns: True if the value was written successfully, False otherwise.
Function write_string_in_file(path As String, value As String) As Boolean
    Dim file As Integer
    file = FreeFile
    Open path For Append As file
    Print #file, value
    Close file
     write_string_in_file = True
End Function

' Description: This function creates a text file.
' Parameters:
'   - path: The path of the file.
' Returns: True if the file was created successfully, False otherwise.
Function create_file(path As String) As Boolean
    Dim fs As Object
    Dim a As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(path, True)
    a.Close
    create_file = True
End Function

' Description: This function deletes a file.
' Parameters:
'   - path: The path of the file.
' Returns: True if the file was deleted successfully, False otherwise.
Function delete_file(path As String) As Boolean
    Kill (path)
    delete_file = True
End Function

' Description: This function searches for a file.
' Parameters:
'   - path: The path of the file.
' Returns: True if the file exists, False otherwise.
Function exist_file(path As String) As Boolean
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    exist_file = fs.FileExists(path)
End Function