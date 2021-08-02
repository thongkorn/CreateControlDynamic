' / ---------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Microsoft Visual Basic .NET (2010) + MS Access 2010+
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / ---------------------------------------------------------------
Imports System.Data.OleDb

Module modDataBase
    Public Conn As OleDbConnection
    Public Cmd As OleDbCommand
    Public DS As DataSet
    Public DR As OleDbDataReader
    Public DA As OleDbDataAdapter
    Public DT As DataTable
    Public strSQL As String     '// Major SQL Statement
    Public strStmt As String    '// Minor SQL Statement
    Public strPath As String = MyPath(Application.StartupPath)
    Public strPathImages As String = strPath & "Images\"

    '// Connect MS Access DataBase
    Public Function MyDBModule() As OleDb.OleDbConnection
        '// Bad Practise because don't check error if have not exist path to DataBase.
        Return New OleDb.OleDbConnection( _
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & "data\dbFood.accdb;Persist Security Info=True")
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    ' / AppPath = C:\My Project\bin\debug
    ' / Replace "\bin\debug" with "\"
    ' / Return : C:\My Project\
    Function MyPath(ByVal AppPath As String) As String
        '/ MessageBox.Show(AppPath);
        AppPath = AppPath.ToLower()
        '/ Return Value
        MyPath = AppPath.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '// If not found folder then put the \ (BackSlash) at the end.
        If Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function

End Module
