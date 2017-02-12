Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports System.Net

Public Class Form1

    Dim Hashes = New String() {"F82EE297EA7487C6461F7C3E3DAF9FC7", "726ED4F381AE412605969EBC85A952A", "4436814B21BE96D1D22B39D74A4D2456", "6DC3168F0E26DF44D678EBF15292DF", "8A557CF0463940C63FE0577B95E5AF3", "C9F9A7B1AD490EAA9FA8B66C6E3E7EE", "BB9F98EDD2928A4C37F4470C0E12DA6", "4F63FBAFAE788E445C67735F2BEDC22A", "40970BD7AA7EF50771822E96D7CF4", "575E8FB6CA584AA24589E2391F805313", "CD164BE93EC0BDA2AA2DFF1DACF01817", "A256BF163F8A7950CDF96E62C6C1517", "DDE8115E1B666E26C9C365F98F9F4D", "A2215B6BC57DD62679BF72B07A93E6", "49209AD7A1B0D9D9793BDECF123F4B4A", "619D606A74EEEBE386D55545AEB92C4", "2C1DBF7C0EA2E5BEBA50C33FF1F147"}
    Dim DownloadLinks = New String() {"http://www.mediafire.com/file/ewoid353twuignn/Base.wz",
        "http://www.mediafire.com/file/abfki7v7fe8q93e/Character.wz",
        "http://www.mediafire.com/file/4134xduecl3x465/Effect.wz",
        "http://www.mediafire.com/file/g1y8p8ke4h71k3w/Etc.wz",
        "http://www.mediafire.com/file/4j0txnqbk3wgr2m/Item.wz",
        "http://www.mediafire.com/file/ye5feg5ebzve7xb/List.wz",
        "http://www.mediafire.com/file/855xf79dzvjs4oo/Map.wz",
        "http://www.mediafire.com/file/kjngug9mup3z6yi/Mob.wz",
        "http://www.mediafire.com/file/910wdb46oddao2u/Morph.wz",
        "http://www.mediafire.com/file/lqsllhy77g9ch0q/Npc.wz",
        "http://www.mediafire.com/file/dl4byldoojy6ya1/Quest.wz",
        "http://www.mediafire.com/file/qxub6kl79xc12x5/Reactor.wz",
        "http://www.mediafire.com/file/uuut8yb76472mu3/Skill.wz",
        "http://www.mediafire.com/file/rjr8rduu9f9hmlb/Sound.wz",
        "http://www.mediafire.com/file/6evvzq5c4uo36f6/String.wz",
        "http://www.mediafire.com/file/6goyfqtjucqs26c/TamingMob.wz",
        "http://www.mediafire.com/file/8224o4mptgxt1ps/UI.wz"}
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
            MsgBox("Missing or too many wz files, verify them and try again.")
            Exit Sub
        End If

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

        If incorrectHashes.Count > 0 Then
            Dim box As Integer = MessageBox.Show("You have corrupt/modified WZ files, would you like to download the clean versions?", "ArcticTools - Parco's Hash Checker", MessageBoxButtons.YesNo)
            If box = DialogResult.Yes Then
                ProgressBar1.Visible = True
                DownloadFiles()
                MsgBox("File downloads complete. If you continue to have issues contact staff for additional support.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
            Else
                MsgBox("You must fix your corrupt files in order to play ArcticMS. Please redownload or if you have used modified WZ files check they are working.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
            End If
        Else
            MsgBox("All of your WZ files are perfect, contact staff for additional support.", MessageBoxButtons.OK, "ArcticTools - Parco's Hash Checker")
        End If
    End Sub

    Private Sub DownloadFiles()
        For Each number In IncorrectHashes
            Dim link = DownloadURL & FileNames(number)
            Dim path = TextBox1.Text & "\" & FileNames(number)
            Dim client As WebClient = New WebClient
            AddHandler client.DownloadProgressChanged, AddressOf client_ProgressChanged
            AddHandler client.DownloadFileCompleted, AddressOf client_DownloadCompleted
            Console.WriteLine("Link:" & link)
            Console.WriteLine("Path:" & path)
            client.DownloadFile(New Uri(link), TextBox1.Text & "\" & FileNames(number))
        Next
    End Sub

    Private Sub client_ProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs)
        Dim bytesIn As Double = Double.Parse(e.BytesReceived.ToString())
        Dim totalBytes As Double = Long.Parse(e.TotalBytesToReceive.ToString())
        Dim percentage As Double = bytesIn / totalBytes * 100
        Console.WriteLine(percentage & " " & totalBytes & " " & bytesIn)
        ProgressBar1.Value = Int32.Parse(Math.Truncate(percentage).ToString())
    End Sub

    Private Sub client_DownloadCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        MessageBox.Show("File downloaded.")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim link = "http://arcticmseu.ddns.net/wz/Character.wz"
        Dim client As WebClient = New WebClient
        Dim path = "C:\Users\Kevin\Desktop\Character.wz"
        Dim attributes As FileAttributes
        attributes = File.GetAttributes(path)
        If ((attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly) Then
            attributes = RemoveAttribute(attributes, FileAttributes.ReadOnly)
            File.SetAttributes(path, attributes)
            Console.WriteLine("File is no longer read only.")
        End If

        client.DownloadFileAsync(New Uri(link), "C:\Users\Kevin\Desktop\Character.wz")
    End Sub

    Public Shared Function RemoveAttribute(ByVal attributes As FileAttributes, ByVal attributesToRemove As FileAttributes) As FileAttributes
        Return attributes And (Not attributesToRemove)
    End Function
End Class
