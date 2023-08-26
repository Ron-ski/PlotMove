Imports System.Reflection
Imports System.Runtime.Remoting.Contexts
Imports System.Text.RegularExpressions

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Program is run from main plot mover GUI
        'A single argument should be passed in the following form
        'D:\Plots1/Drive Bay 1/C:\Plot Test\plot-k32-c07-2023-06-10-16-22-5-PlotTest.mov02
        'Destination directory, nickname and the full plot path with extension.movXX where XX is the drive number
        'A forward slash / must be used to seperate the arguments.

        Dim MyDebug = False   'Set to true for testing purposes
        Dim MyDebugArguments = "C:\Plot Test/Drive Bay 1/D:\Plot Test\plot-k32-c07-2023-06-10-16-22-5-PlotTest.mov99"

        Dim PlotName As String
        Dim DestDir As String


        Dim RecievedText As String = ""
        Dim NickName As String
        Dim MyArguments() As String = {"", "", "", ""}
        Dim DriveNumber As String = ""

        If Command() = "" And MyDebug = False Then
            MsgBox("No arguments passed, the program will close.")

        Else


            Try
                RecievedText = Command()

                If MyDebug Then RecievedText = MyDebugArguments

                MyArguments = GetArguments(RecievedText)

                DestDir = MyArguments(0)                             'Get the destination directory
                NickName = MyArguments(1)                            'Get the nickname
                PlotName = MyArguments(2)                            'Get the path and file name we are moving
                DriveNumber = MyArguments(3)                         'Get the drive number


                'MsgBox("Disable messages for debug purposes")

                Dim SizeinByte As Long = FileLen(PlotName)
                Dim SizeinGB As String = (Math.Round(SizeinByte / 1024 / 1024 / 1024, 2)).ToString()


                Dim FileName As String = System.IO.Path.GetFileName(PlotName)                          'Get just the filename
                Dim NewPlotName As String = FileName.Substring(0, FileName.Length - 5) & "plot"        'Change the extension for later use


                'Construct the logfile name

                Dim MyPath As String = My.Application.Info.DirectoryPath & "\Logs"
                Dim MyLogFile As String = MyPath & "\" & System.DateTime.Now.ToString("yyyy-MM-dd") & " " & NickName & ", Drive" & DriveNumber & ".log"

                'Create the start time for the logfile

                Dim StartTime As String = Now().ToString
                Dim LogText As String = StartTime & ", " & DestDir & "," & PlotName & ", " & SizeinGB & ", "

                'Move the files - shows the standard windows dialog

                My.Computer.FileSystem.MoveFile(PlotName, DestDir & "\" & FileName, FileIO.UIOption.AllDialogs)

                'Rename the file back to .Plot

                My.Computer.FileSystem.RenameFile(DestDir & "\" & FileName, NewPlotName)


                'Calculate the duration

                Dim dFrom As DateTime
                Dim dTo As DateTime
                Dim EndTime As String = Now().ToString

                Dim Duration As String = "Not calculated"

                If DateTime.TryParse(StartTime, dFrom) AndAlso DateTime.TryParse(EndTime, dTo) Then
                    Dim TS As TimeSpan = dTo - dFrom
                    Dim hour As Integer = TS.Hours
                    Dim mins As Integer = TS.Minutes
                    Dim secs As Integer = TS.Seconds
                    Duration = ((hour.ToString("00") & ":") + mins.ToString("00") & ":") + secs.ToString("00")

                End If

                'Check if the Logs directory exists and if it doesn't create it

                If Not System.IO.Directory.Exists(MyPath) Then
                    System.IO.Directory.CreateDirectory(MyPath)
                End If


                ' Open the log file and write to it.

                Dim file As System.IO.StreamWriter
                file = My.Computer.FileSystem.OpenTextFileWriter(MyLogFile, True)
                file.WriteLine(LogText & Duration)
                file.Close()

            Catch ex As Exception

                If MyDebug Then MsgBox(ex.Message & vbCrLf & vbCrLf & RecievedText)

                Dim MyPath As String = My.Application.Info.DirectoryPath & "\Logs"

                If Not System.IO.Directory.Exists(MyPath) Then
                    System.IO.Directory.CreateDirectory(MyPath)
                End If


                Dim file As System.IO.StreamWriter
                file = My.Computer.FileSystem.OpenTextFileWriter(MyPath & "\Error.log", True)
                file.WriteLine(Now() & " The following error was generated:" & vbCrLf & ex.Message & vbCrLf & "The passed arguments are: " & Command())
                file.Close()

            End Try

            'If the user cancelled or there was an error we need to rename the .movX file back to .plot
            'With the above being in a try catch block it causes issues and I don't know the correct way to deal with it.

            RecievedText = Command()

            If MyDebug Then RecievedText = MyDebugArguments


            MyArguments = GetArguments(RecievedText)

            DestDir = MyArguments(0)                             'Get the destination directory
            NickName = MyArguments(1)                            'Get the nickname
            PlotName = MyArguments(2)                            'Get the path and file name we are moving


            If (System.IO.File.Exists(PlotName)) Then

                If MyDebug Then MsgBox("file still there and has not been moved")

                Dim FileName As String = System.IO.Path.GetFileName(PlotName)                          'Get just the filename
                Dim NewPlotName As String = FileName.Substring(0, FileName.Length - 5) & "plot"        'Change the extension for later use
                My.Computer.FileSystem.RenameFile(PlotName, NewPlotName)                               'Rename the file back to .plot 

            End If
        End If


        Close()


    End Sub

    Private Function GetArguments(ByVal RecievedText As String) As String()

        Dim myResult(3) As String  '0 = destdir 1 = Nickname  2 = Path & plot name 3 = Drive Number
        Dim MyLength As Integer
        Dim FirstForwardSlash

        MyLength = RecievedText.Length
        FirstForwardSlash = RecievedText.IndexOf("/")


        myResult(0) = RecievedText.Substring(0, FirstForwardSlash)                             'Get the destination directory


        Dim NickNamePlotName As String = RecievedText.Substring(FirstForwardSlash + 1, MyLength - FirstForwardSlash - 1)    'Get the path and file name we are moving

        MyLength = NickNamePlotName.Length
        FirstForwardSlash = NickNamePlotName.IndexOf("/")

        myResult(1) = NickNamePlotName.Substring(0, FirstForwardSlash)                                       'Get the destination directory

        myResult(2) = NickNamePlotName.Substring(FirstForwardSlash + 1, MyLength - FirstForwardSlash - 1)    'Get the path and file name we are moving

        myResult(3) = myResult(2).Substring(myResult(2).Length - 2, 2)                                       'Get the drive number

        Return myResult


    End Function
End Class
