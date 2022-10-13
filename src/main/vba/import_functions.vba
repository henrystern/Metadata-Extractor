Sub activateHyperlinks()
' if the hyperlinks were not imported into excel as a formula this will activate the links
    Dim Table_Name As String
    Table_Name = "PDF_Metadata" 'replace PDF_Metadata with the name of your table
    
    Dim Column_Header As String
    Column_Header = "File Path" 'replace File Path with name of column

    Dim c As Range
    For Each c In Range(Table_Name & "[" & Column_Header & "]")
        c.Value = c.Value
    Next c
End Sub

Sub shortenHyperlinks()
' If hyperlinks exceed the excel formula length limit they will not work
' Steps to fix:
'   1. regenerate the metadata spreadsheet with relative path and no hyperlinks
'   2. run this command with the appropriate inputs highlighted below
' it works by shortening the link to the DOS path
' only works with relative links but could be modified pretty easily
    
    Dim root_path As String
    ' INPUT REQUIRED
    root_path = "C:\Users\User\Desktop\" 'replace with common root of pdf locations
    
    Dim Table_Name As String
    ' INPUT REQUIRED
    Table_Name = "PDF_Metadata" 'replace PDF_Metadata with the name of your table
    
    Dim Column_Header As String
    ' INPUT REQUIRED
    Column_Header = "File Path" 'replace File Path with name of column

    Dim i As Integer 'to count rows for reporting
    i = 2
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim root_path_obj As Object
    Set rel_path_obj = FSO.GetFolder(root_path)
            
    Dim short_root_path As String
    short_root_path = rel_path_obj.shortpath
    
    Dim short_root_path_length As Integer
    short_root_path_length = Len(short_root_path) + 1 '+1 so the relative short path excludes the first \
    
    Dim c As Range
    For Each c In Range(Table_Name & "[" & Column_Header & "]")
        On Error Resume Next
        
        If Len(c.Value) > 20 Then 'if path is longer than this the link wont work 
                                  'if long shorten to unreadable
            Dim file_path As String
            file_path = c.Value
            
            Dim full_path As String
            full_path = root_path & Left(file_path, InStrRev(file_path, "\"))
                       
            Dim file_name As String
            file_name = Right(file_path, Len(file_path) - InStrRev(file_path, "\"))
            
            Dim full_path_obj As Object
            Set full_path_obj = FSO.GetFolder(full_path)
    
            Dim short_path As String
            short_path = full_path_obj.shortpath

            Dim rel_short_path As String
            rel_short_path = Right(short_path, Len(short_path) - short_root_path_length) & "\" & file_name 'relative short path

            If Len(rel_short_path) > 204 Then 'if still too long shorten file name as well as path
                                              ' (makes unreadable) 
                                              ' doesnt work yet because the file path is too long
                                              ' I tried guessing the dos format but it treats spaces strangely
                                              ' and broke most links
                                              ' probably a fringe scenario so didn't bother fixing
                
                rel_short_path = Right(file_short_path, Len(file_short_path) - short_root_path_length)
                Debug.Print rel_short_path

                If Len(rel_short_path) > 204 Then
                    rel_short_path = "[File Name Too Long] " & rel_short_path 'if still too long then 
                                                                              'it is unfixable without
                                                                              ' renaming the files 
                                                                              ' -- mark as broken
                End If
            End If
            
            c.Value = "=HYPERLINK(" & Chr(34) & rel_short_path & Chr(34) & ")"
            Debug.Print "Shortened path at row " & i
            
        Else
            c.Value = "=HYPERLINK(" & Chr(34) & c.Value & Chr(34) & ")" 'if not long 
                                                                        'keep the links 
                                                                        'the same so that 
                                                                        'they are more readable
        End If
        
        i = i + 1
    Next c
End Sub