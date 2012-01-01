Attribute VB_Name = "myloopbase"
Sub myloopbase()

    Const dest_sheetname = "dest"
    Const dest_col_st = "A"
    Const dest_line_st = "2"

    Const db1_sheetname = "db1"
    Const db1_col_st = "A"
    Const db1_line_st = "1"

    Const msg_fin = "FINISHED"
    Const fileinfo = "csv file,*.csv"

    Dim self_sheetname As String
    Dim self_exls, dest_exls, db1_exls As Object

    Dim dest_line As Long
    Dim db1_line, db1_line_max As Long

    ' universal variable
    Dim flag As Boolean
    Dim filepath As String

    ' OFF the automatic calculation
    With Application
        .Calculation = xlManual
    End With

    ' set the active sheet name
    self_sheetname = ActiveSheet.Name

    ' set self_exls object
    Set self_exls = ActiveWorkbook.Sheets(self_sheetname)

    ' set and clean dest_exls object
    Set dest_exls = ActiveWorkbook.Sheets(dest_sheetname)
    dest_line = IIf( _
        dest_exls.Range(dest_col_st & CStr(CInt(dest_line_st) + 1)) <> "", _
        dest_exls.Range(dest_col_st & dest_line_st).End(xlDown).Row, 1 _
    )
    dest_exls.Range( _
        dest_line_st & ":" & CStr(CInt(dest_line_st) + dest_line - 1) _
    ).Delete

    ' set database object
    MsgBox "Choose a database file."
    ' open a chosen database
    flag = GetFileName.get_filepath(filepath, fileinfo)
    If flag = False Then
        MsgBox "Can't get a database file."
        Exit Sub
    End If
    Workbooks.Open Filename:=filepath
    Set db1_exls = ActiveWorkbook.Sheets(db1_sheetname)

    ' main loop
    db1_line_max = db1_exls.Range(db1_col_st & db1_line_st).End(xlDown).Row
    For db1_line = CInt(db1_line_st) To db1_line_max
        ' update statusbar to see a computing
        Application.StatusBar = "Please, wait a minutes. Loop: " & _
                                    CStr(db1_line) & "/" & CStr(db1_line_max)

        ' initialize

        ' update

        ' for next


    Next db1_line

    ' close database object without a confirm
    db1_exls.Activate
    Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False

    ' activate self_exls
    self_exls.Activate

    ' ON the automatic calculation
    With Application
        .Calculation = xlAutomatic
    End With

    ' print the message
    MsgBox msg_fin

End Sub
