Attribute VB_Name = "GetFilePath"
' ---------------------------------------------------------------------------
' Name: get_filepath
' Description:
'   You can choose a file on the file dialog.
'   Then, You can get a name of selected file.
' Arguments:
'   filepath
'   fileinfo
' Return
'   True
'   False
' ---------------------------------------------------------------------------
Public Function get_filepath( _
    ByRef filepath As String, _
    ByVal fileinfo As String _
) As Boolean

    Dim self_filepath As String

    ' get my file path and change active directory
    self_filepath = ActiveWorkbook.FullName
    ChDir(Replace(self_filepath, Dir(self_filepath), ""))

    ' open the file dialog
    filepath = Application.GetOpenFilename(fileinfo)

    ' when you canceled or the file path is NOT
    If filepath = "False" Or Dir(filepath) = "" Then

        get_filepath = False
        Exit Function

    End If

    get_filepath = True

End Function
