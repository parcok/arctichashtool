﻿Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Text.RegularExpressions

Public Class Form1

    Dim client As WebClient = New WebClient
    Dim Hashes = New String() {"F72EE297EA7487C6461F7C3E3DAF9FC7", "726ED4F381AE412605969EBC85A952A", "4436814B21BE96D1D22B39D74A4D2456", "6DC3168F0E26DF44D678EBF15292DF", "8A557CF0463940C63FE0577B95E5AF3", "C9F9A7B1AD490EAA9FA8B66C6E3E7EE", "BB9F98EDD2928A4C37F4470C0E12DA6", "4F63FBAFAE788E445C67735F2BEDC22A", "40970BD7AA7EF50771822E96D7CF4", "575E8FB6CA584AA24589E2391F805313", "CD164BE93EC0BDA2AA2DFF1DACF01817", "A256BF163F8A7950CDF96E62C6C1517", "DDE8115E1B666E26C9C365F98F9F4D", "A2215B6BC57DD62679BF72B07A93E6", "49209AD7A1B0D9D9793BDECF123F4B4A", "619D606A74EEEBE386D55545AEB92C4", "2C1DBF7C0EA2E5BEBA50C33FF1F147"}
    Dim fileHost = "http://arcticmseu.ddns.net/wz/"
    Dim FileNames = New String() {"Base.wz", "Character.wz", "Effect.wz", "Etc.wz", "Item.wz", "List.wz", "Map.wz", "Mob.wz", "Morph.wz", "Npc.wz", "Quest.wz", "Reactor.wz", "Skill.wz", "Sound.wz", "String.wz", "TamingMob.wz", "UI.wz"}
    Dim UserHashes(16) As String
    Dim IncorrectHashes As New ArrayList()

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim fInfo As FileInfo()
        Dim i As Integer = 0
        Dim path As String = TextBox1.Text
        Dim dInfo As New DirectoryInfo(path)

        fInfo = dInfo.GetFiles("*.txt")
        Dim files As String()
        Dim FilePath As String
        files = Directory.GetFiles(path, "*.wz")

        If files.Count <> 17 Then
            MsgBox("Missing or too many wz files, verify them and try again.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
            Exit Sub
        End If

        MsgBox("Checking Wz files, this may take a minute, let the program run.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
        Dim counter = 0
        For Each FilePath In files
            ' Get the file attributes and remove readonly
            Dim attributes As FileAttributes
            attributes = File.GetAttributes(FilePath)
            If ((attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly) Then
                attributes = RemoveAttribute(attributes, FileAttributes.ReadOnly)
                File.SetAttributes(FilePath, attributes)
            End If
            'Do Something With The File
            Using RD As New FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
                'Dim RD As FileStream = New FileStream(File, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
                '   RD = New FileStream(File, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
                Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
                md5.ComputeHash(RD)
                RD.Dispose()
                RD.Close()

                'converting the bytes into string
                Dim hash As Byte() = md5.Hash
                Dim SB As StringBuilder = New StringBuilder
                Dim HB As Byte
                For Each HB In hash
                    SB.Append(String.Format("{0:X1}", HB))
                Next
                UserHashes(counter) = SB.ToString
                counter = counter + 1
            End Using
            ProgressBar1.Value = 100 / files.Count * counter
        Next

        ' Comparing hashes 
        counter = 0
        For Each hash As String In UserHashes
            If Hashes(counter) <> UserHashes(counter) Then
                IncorrectHashes.Add(counter)
            End If
            counter = counter + 1
        Next

        For Each fileName In IncorrectHashes
            Console.WriteLine(fileName)
        Next
        If IncorrectHashes.Count > 0 Then
            Dim box As Integer = MessageBox.Show("You have corrupt/modified WZ files, would you like to download the clean versions?", "ArcticTools - Parco's Hash Checker", MessageBoxButtons.YesNo)
            If box = DialogResult.Yes Then
                DownloadFiles()
            Else
                MsgBox("You must fix your corrupt files in order to play ArcticMS. Please redownload or if you have used modified WZ files check they are working.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
            End If
        Else
            MsgBox("All of your WZ files are perfect, contact staff for additional support.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
        End If
    End Sub

    Private Sub DownloadFiles()
        'For Each number In IncorrectHashes
        Dim link = fileHost & FileNames(IncorrectHashes(0))
        Console.WriteLine(link)
        Dim path = TextBox1.Text & "\" & FileNames(IncorrectHashes(0))
        AddHandler client.DownloadProgressChanged, AddressOf client_ProgressChanged
            AddHandler client.DownloadFileCompleted, AddressOf client_DownloadCompleted
            Console.WriteLine("Link:" & link)
            Console.WriteLine("Path:" & path)
        client.DownloadFileAsync(New Uri(link), TextBox1.Text & "\" & FileNames(IncorrectHashes(0)))
        'Next
        MsgBox("File downloads complete. If you continue to have issues contact staff for additional support.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
    End Sub

    Private Sub client_ProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs)
        Dim bytesIn As Double = Double.Parse(e.BytesReceived.ToString())
        Dim totalBytes As Double = Long.Parse(e.TotalBytesToReceive.ToString())
        Dim percentage As Double = bytesIn / totalBytes * 100
        'Console.WriteLine("Percentage: " & percentage & " Total Bytes: " & totalBytes & " Bytes In: " & bytesIn)
        ProgressBar1.Value = Int32.Parse(Math.Truncate(percentage).ToString())
    End Sub

    Private Sub client_DownloadCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        MsgBox("File downloaded.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
        If IncorrectHashes.Count > 0 Then
            IncorrectHashes.RemoveAt(0)
            DownloadFiles()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim link = "http://arcticmseu.ddns.net/wz/Etc.wz"
        Dim client As WebClient = New WebClient
        Dim path = "C:\Users\Kevin\Desktop\Etc.wz"
        Dim attributes As FileAttributes
        attributes = File.GetAttributes(path)
        If ((attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly) Then
            attributes = RemoveAttribute(attributes, FileAttributes.ReadOnly)
            File.SetAttributes(path, attributes)
        End If

        AddHandler client.DownloadProgressChanged, AddressOf client_ProgressChanged
        AddHandler client.DownloadFileCompleted, AddressOf client_DownloadCompleted
        client.DownloadFileAsync(New Uri(link), "C:\Users\Kevin\Desktop\Character.wz")
        'isolateMediafireLink("http://www.mediafire.com/file/abfki7v7fe8q93e/Character.wz")
    End Sub

    Public Shared Function RemoveAttribute(ByVal attributes As FileAttributes, ByVal attributesToRemove As FileAttributes) As FileAttributes
        Return attributes And (Not attributesToRemove)
    End Function

    Public Shared Function isolateMediafireLink(mediafireLink As String)
        Dim webc As WebClient = New WebClient
        webc.Headers("User-Agent") = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36"
        webc.Headers("Accept-Language") = "en-US,en;q=0.8"
        webc.Headers("Accept") = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        Dim webSource As String = webc.DownloadString(mediafireLink)
        'Console.WriteLine(webSource)
        Dim regex = New Regex("http:\/\/download\d{4}\.mediafire.com.+?'")
        Dim match As Match = regex.Match(webSource)
        Dim directDownload As String
        If match.Success Then
            directDownload = match.Groups(1).Value
            Console.WriteLine("Link: " & directDownload)
            MsgBox(directDownload, MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
        Else
            Console.WriteLine("No value found.")
            MsgBox("No value found.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
        End If
        Return "Hello"
    End Function
End Class
