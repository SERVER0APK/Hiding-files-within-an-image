Imports DevExpress.XtraEditors
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Partial Public Class Hide_files

    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub tileBar_SelectedItemChanged(ByVal sender As Object, ByVal e As TileItemEventArgs) Handles tileBar.SelectedItemChanged
        navigationFrame.SelectedPageIndex = tileBarGroupTables.Items.IndexOf(e.Item)
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click

        Dim openFileDialog1 As New OpenFileDialog()
        Dim openFileDialog2 As New OpenFileDialog()
        Dim saveFileDialog1 As New SaveFileDialog()

        ' Set the filter to JPG files
        openFileDialog1.Filter = "JPG Files (*.jpg)|*.jpg"
        openFileDialog2.Filter = "JPG Files (*.jpg)|*.jpg"
        saveFileDialog1.Filter = "JPG Files (*.jpg)|*.jpg"

        ' Show the open file dialog
        If openFileDialog1.ShowDialog() = DialogResult.OK AndAlso openFileDialog2.ShowDialog() = DialogResult.OK Then
            ' Read the input file bytes
            Dim inputFileBytes1 As Byte() = File.ReadAllBytes(openFileDialog1.FileName)
            Dim inputFileBytes2 As Byte() = File.ReadAllBytes(openFileDialog2.FileName)

            ' Get the end index of the first image
            Dim imageEndIndex As Integer = -1
            Dim imageEndSequence As Byte() = {&HFF, &HD9}
            For i As Integer = inputFileBytes1.Length - imageEndSequence.Length To 0 Step -1
                Dim found As Boolean = True
                For j As Integer = 0 To imageEndSequence.Length - 1
                    If inputFileBytes1(i + j) <> imageEndSequence(j) Then
                        found = False
                        Exit For
                    End If
                Next
                If found Then
                    imageEndIndex = i + imageEndSequence.Length
                    Exit For
                End If
            Next

            If imageEndIndex = -1 Then
                Console.WriteLine("End of image not found in file 1")
                Exit Sub
            End If

            ' Append the second image bytes to the first image bytes
            Dim outputFileBytes(imageEndIndex + inputFileBytes2.Length - 1) As Byte
            Array.Copy(inputFileBytes1, outputFileBytes, imageEndIndex)
            Array.Copy(inputFileBytes2, 0, outputFileBytes, imageEndIndex, inputFileBytes2.Length)

            ' Show the save file dialog
            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                ' Save the output file
                File.WriteAllBytes(saveFileDialog1.FileName, outputFileBytes)
            End If
        End If

    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        Dim inputFileBytes As Byte()
        Dim outputDirectory As String = "C:\Users\alicr\OneDrive\Desktop\New folder (2)\output\New folder"

        Dim openFileDialog1 As New OpenFileDialog()

        With openFileDialog1
            .Filter = "JPG Files (*.jpg)|*.jpg"
            .Title = "Select a JPEG file"
            .Multiselect = False
        End With

        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            inputFileBytes = File.ReadAllBytes(openFileDialog1.FileName)
            outputDirectory = Path.GetDirectoryName(openFileDialog1.FileName)
        End If

        Dim imageStartSequence As Byte() = {&HFF, &HD8, &HFF, &HE0}

        Dim startIndex As Integer = 0
        Dim endIndex As Integer = 0
        Dim imageCount As Integer = 0

        Do
            Dim imageStartIndex As Integer = -1
            Dim imageEndIndex As Integer = -1

            For i As Integer = startIndex To inputFileBytes.Length - imageStartSequence.Length
                Dim found As Boolean = True
                For j As Integer = 0 To imageStartSequence.Length - 1
                    If inputFileBytes(i + j) <> imageStartSequence(j) Then
                        found = False
                        Exit For
                    End If
                Next

                If found Then
                    imageStartIndex = i
                    Exit For
                End If
            Next

            If imageStartIndex <> -1 Then
                For i As Integer = imageStartIndex + imageStartSequence.Length To inputFileBytes.Length - imageStartSequence.Length
                    Dim found As Boolean = True
                    For j As Integer = 0 To imageStartSequence.Length - 1
                        If inputFileBytes(i + j) <> imageStartSequence(j) Then
                            found = False
                            Exit For
                        End If
                    Next

                    If found Then
                        imageEndIndex = i - 1
                        Exit For
                    End If
                Next

                If imageEndIndex <> -1 Then
                    Dim imageLength As Integer = imageEndIndex - imageStartIndex + 1
                    Dim imageData(imageLength - 1) As Byte
                    Buffer.BlockCopy(inputFileBytes, imageStartIndex, imageData, 0, imageLength)

                    ' Add missing FF value at the start of the image data
                    imageData(0) = 255

                    Dim outputFilePath As String = Path.Combine(outputDirectory, "output_" & imageCount & ".jpg")
                    File.WriteAllBytes(outputFilePath, imageData)

                    startIndex = imageEndIndex + 1
                    imageCount += 1
                Else
                    ' Save the last image
                    Dim imageLength As Integer = inputFileBytes.Length - imageStartIndex
                    Dim imageData(imageLength - 1) As Byte
                    Buffer.BlockCopy(inputFileBytes, imageStartIndex, imageData, 0, imageLength)

                    ' Add missing FF value at the start of the image data
                    imageData(0) = 255

                    Dim outputFilePath As String = Path.Combine(outputDirectory, "output_" & imageCount & ".jpg")
                    File.WriteAllBytes(outputFilePath, imageData)

                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
    End Sub

    Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        Dim imageFilePath As String = ""
        Dim pdfFilePath As String = ""
        Dim outputFilePath As String = ""
        ' Open the image file dialog
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            imageFilePath = OpenFileDialog1.FileName
            ' Open the PDF file dialog
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                pdfFilePath = OpenFileDialog1.FileName
                ' Check if the output file already exists
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    outputFilePath = SaveFileDialog1.FileName
                    ' Check if the output file already exists
                    If File.Exists(outputFilePath) Then
                        Dim result As DialogResult = MessageBox.Show("The output file already exists. Do you want to replace it?", "File already exists", MessageBoxButtons.YesNo)
                        If result = DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                    ' Read the image and PDF files as byte arrays
                    Dim imageBytes As Byte() = File.ReadAllBytes(imageFilePath)
                    Dim pdfBytes As Byte() = File.ReadAllBytes(pdfFilePath)
                    ' Find the end of the image in the image byte array
                    Dim imageEndIndex As Integer = -1
                    Dim imageEndSequence As Byte() = {&HFF, &HD9}
                    For i As Integer = imageBytes.Length - imageEndSequence.Length To 0 Step -1
                        Dim found As Boolean = True
                        For j As Integer = 0 To imageEndSequence.Length - 1
                            If imageBytes(i + j) <> imageEndSequence(j) Then
                                found = False
                                Exit For
                            End If
                        Next
                        If found Then
                            imageEndIndex = i + imageEndSequence.Length
                            Exit For
                        End If
                    Next
                    If imageEndIndex = -1 Then
                        MsgBox("End of image not found")
                        Exit Sub
                    End If
                    ' Create a new byte array to hold the merged image and PDF
                    Dim mergedBytes(imageEndIndex + pdfBytes.Length - 1) As Byte
                    ' Copy the image bytes to the merged byte array
                    Array.Copy(imageBytes, mergedBytes, imageEndIndex)
                    ' Copy the PDF bytes to the merged byte array
                    Array.Copy(pdfBytes, 0, mergedBytes, imageEndIndex, pdfBytes.Length)
                    ' Write the merged byte array to a new image file
                    File.WriteAllBytes(outputFilePath, mergedBytes)
                End If
            End If
        End If
    End Sub

    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|PDF Files (*.pdf)|*.pdf"
        If openFileDialog1.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If
        Dim inputFilePath As String = openFileDialog1.FileName
        Dim inputFileBytes As Byte() = File.ReadAllBytes(inputFilePath)
        Dim pdfStartSequence As Byte() = {&H25, &H50, &H44, &H46}
        Dim pdfStartIndex As Integer = -1
        For i As Integer = 0 To inputFileBytes.Length - pdfStartSequence.Length
            Dim found As Boolean = True
            For j As Integer = 0 To pdfStartSequence.Length - 1
                If inputFileBytes(i + j) <> pdfStartSequence(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then
                pdfStartIndex = i
                Exit For
            End If
        Next
        If pdfStartIndex = -1 Then
            Console.WriteLine("PDF not found in image")
            Exit Sub
        End If
        Dim pdfBytes(inputFileBytes.Length - pdfStartIndex - 1) As Byte
        Array.Copy(inputFileBytes, pdfStartIndex, pdfBytes, 0, pdfBytes.Length)
        Dim SaveFileDialog2 As New SaveFileDialog()
        SaveFileDialog2.Filter = "PDF Files (*.pdf)|*.pdf"
        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog2.FileName
            File.WriteAllBytes(outputFilePath, pdfBytes)
        End If
        Dim imageBytes(inputFileBytes.Length - pdfBytes.Length - 1) As Byte
        Array.Copy(inputFileBytes, 0, imageBytes, 0, pdfStartIndex)
        Dim SaveFileDialog3 As New SaveFileDialog()
        SaveFileDialog3.Filter = "JPEG Files (*.jpg)|*.jpg"
        If SaveFileDialog3.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog3.FileName
            File.WriteAllBytes(outputFilePath, imageBytes)
        End If
    End Sub

    Private Sub SimpleButton6_Click(sender As Object, e As EventArgs) Handles SimpleButton6.Click
        Dim imageFilePath As String = ""
        Dim pdfFilePath As String = ""
        Dim outputFilePath As String = ""

        OpenFileDialog1.Filter = "Image Files|*.jpg;*.jpeg|All Files|*.*"

        ' Open the image file dialog
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            imageFilePath = OpenFileDialog1.FileName

            OpenFileDialog1.Filter = "PDF Files|*.docx|All Files|*.*"

            ' Open the PDF file dialog
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                pdfFilePath = OpenFileDialog1.FileName

                SaveFileDialog1.Filter = "JPEG Files|*.jpg|All Files|*.*"

                ' Check if the output file already exists
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    outputFilePath = Path.ChangeExtension(SaveFileDialog1.FileName, ".jpg")

                    ' Check if the output file already exists
                    If File.Exists(outputFilePath) Then
                        Dim result As DialogResult = MessageBox.Show("The output file already exists. Do you want to replace it?", "File already exists", MessageBoxButtons.YesNo)
                        If result = DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    ' Read the image and PDF files as byte arrays
                    Dim imageBytes As Byte() = File.ReadAllBytes(imageFilePath)
                    Dim pdfBytes As Byte() = File.ReadAllBytes(pdfFilePath)

                    ' Find the end of the image in the image byte array
                    Dim imageEndIndex As Integer = -1
                    Dim imageEndSequence As Byte() = {&HFF, &HD9}
                    For i As Integer = imageBytes.Length - imageEndSequence.Length To 0 Step -1
                        Dim found As Boolean = True
                        For j As Integer = 0 To imageEndSequence.Length - 1
                            If imageBytes(i + j) <> imageEndSequence(j) Then
                                found = False
                                Exit For
                            End If
                        Next
                        If found Then
                            imageEndIndex = i + imageEndSequence.Length
                            Exit For
                        End If
                    Next

                    If imageEndIndex = -1 Then
                        MsgBox("End of image not found")
                        Exit Sub
                    End If

                    ' Create a new byte array to hold the merged image and PDF
                    Dim mergedBytes(imageEndIndex + pdfBytes.Length - 1) As Byte

                    ' Copy the image bytes to the merged byte array
                    Array.Copy(imageBytes, mergedBytes, imageEndIndex)

                    ' Copy the PDF bytes to the merged byte array
                    Array.Copy(pdfBytes, 0, mergedBytes, imageEndIndex, pdfBytes.Length)

                    ' Write the merged byte array to a new image file
                    File.WriteAllBytes(outputFilePath, mergedBytes)
                End If
            End If
        End If
    End Sub

    Private Sub SimpleButton5_Click(sender As Object, e As EventArgs) Handles SimpleButton5.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|word Files (*.docx)|*.docx"
        If openFileDialog1.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        Dim inputFilePath As String = openFileDialog1.FileName

        Dim inputFileBytes As Byte() = File.ReadAllBytes(inputFilePath)

        Dim pdfStartSequence As Byte() = {&H50, &H4B, &H3, &H4}

        Dim pdfStartIndex As Integer = -1
        For i As Integer = 0 To inputFileBytes.Length - pdfStartSequence.Length
            Dim found As Boolean = True
            For j As Integer = 0 To pdfStartSequence.Length - 1
                If inputFileBytes(i + j) <> pdfStartSequence(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then
                pdfStartIndex = i
                Exit For
            End If
        Next

        If pdfStartIndex = -1 Then
            MsgBox("docx not found in image")
            Exit Sub
        End If

        Dim pdfBytes(inputFileBytes.Length - pdfStartIndex - 1) As Byte
        Array.Copy(inputFileBytes, pdfStartIndex, pdfBytes, 0, pdfBytes.Length)

        Dim SaveFileDialog2 As New SaveFileDialog()
        SaveFileDialog2.Filter = "docx Files (*.docx)|*.docx"

        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog2.FileName
            File.WriteAllBytes(outputFilePath, pdfBytes)
        End If

        Dim imageBytes(inputFileBytes.Length - pdfBytes.Length - 1) As Byte
        Array.Copy(inputFileBytes, 0, imageBytes, 0, pdfStartIndex)

        Dim SaveFileDialog3 As New SaveFileDialog()
        SaveFileDialog3.Filter = "JPEG Files (*.jpg)|*.jpg"

        If SaveFileDialog3.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog3.FileName
            File.WriteAllBytes(outputFilePath, imageBytes)
        End If
    End Sub

    Private Sub SimpleButton7_Click(sender As Object, e As EventArgs) Handles SimpleButton7.Click
        OpenFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|RAR Files (*.RAR)|*.RAR"
        If OpenFileDialog1.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        Dim inputFilePath As String = OpenFileDialog1.FileName

        Dim inputFileBytes As Byte() = File.ReadAllBytes(inputFilePath)

        Dim pdfStartSequence As Byte() = {&H52, &H61, &H72, &H21}

        Dim pdfStartIndex As Integer = -1
        For i As Integer = 0 To inputFileBytes.Length - pdfStartSequence.Length
            Dim found As Boolean = True
            For j As Integer = 0 To pdfStartSequence.Length - 1
                If inputFileBytes(i + j) <> pdfStartSequence(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then
                pdfStartIndex = i
                Exit For
            End If
        Next

        If pdfStartIndex = -1 Then
            MsgBox("RAR not found in image")
            Exit Sub
        End If

        Dim pdfBytes(inputFileBytes.Length - pdfStartIndex - 1) As Byte
        Array.Copy(inputFileBytes, pdfStartIndex, pdfBytes, 0, pdfBytes.Length)

        Dim SaveFileDialog2 As New SaveFileDialog()
        SaveFileDialog2.Filter = "RAR Files (*.RAR)|*.RAR"

        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog2.FileName
            File.WriteAllBytes(outputFilePath, pdfBytes)
        End If

        Dim imageBytes(inputFileBytes.Length - pdfBytes.Length - 1) As Byte
        Array.Copy(inputFileBytes, 0, imageBytes, 0, pdfStartIndex)

        Dim SaveFileDialog3 As New SaveFileDialog()
        SaveFileDialog3.Filter = "JPEG Files (*.jpg)|*.jpg"

        If SaveFileDialog3.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog3.FileName
            File.WriteAllBytes(outputFilePath, imageBytes)
        End If
    End Sub

    Private Sub SimpleButton8_Click(sender As Object, e As EventArgs) Handles SimpleButton8.Click
        Dim imageFilePath As String = ""
        Dim pdfFilePath As String = ""
        Dim outputFilePath As String = ""

        OpenFileDialog1.Filter = "Image Files|*.jpg;*.jpeg|All Files|*.*"

        ' Open the image file dialog
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            imageFilePath = OpenFileDialog1.FileName

            OpenFileDialog1.Filter = "RAR Files|*.rar|All Files|*.*"

            ' Open the PDF file dialog
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                pdfFilePath = OpenFileDialog1.FileName

                SaveFileDialog1.Filter = "JPEG Files|*.jpg|All Files|*.*"

                ' Check if the output file already exists
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    outputFilePath = Path.ChangeExtension(SaveFileDialog1.FileName, ".jpg")

                    ' Check if the output file already exists
                    If File.Exists(outputFilePath) Then
                        Dim result As DialogResult = MessageBox.Show("The output file already exists. Do you want to replace it?", "File already exists", MessageBoxButtons.YesNo)
                        If result = DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    ' Read the image and PDF files as byte arrays
                    Dim imageBytes As Byte() = File.ReadAllBytes(imageFilePath)
                    Dim pdfBytes As Byte() = File.ReadAllBytes(pdfFilePath)

                    ' Find the end of the image in the image byte array
                    Dim imageEndIndex As Integer = -1
                    Dim imageEndSequence As Byte() = {&HFF, &HD9}
                    For i As Integer = imageBytes.Length - imageEndSequence.Length To 0 Step -1
                        Dim found As Boolean = True
                        For j As Integer = 0 To imageEndSequence.Length - 1
                            If imageBytes(i + j) <> imageEndSequence(j) Then
                                found = False
                                Exit For
                            End If
                        Next
                        If found Then
                            imageEndIndex = i + imageEndSequence.Length
                            Exit For
                        End If
                    Next

                    If imageEndIndex = -1 Then
                        MsgBox("End of image not found")
                        Exit Sub
                    End If

                    ' Create a new byte array to hold the merged image and PDF
                    Dim mergedBytes(imageEndIndex + pdfBytes.Length - 1) As Byte

                    ' Copy the image bytes to the merged byte array
                    Array.Copy(imageBytes, mergedBytes, imageEndIndex)

                    ' Copy the PDF bytes to the merged byte array
                    Array.Copy(pdfBytes, 0, mergedBytes, imageEndIndex, pdfBytes.Length)

                    ' Write the merged byte array to a new image file
                    File.WriteAllBytes(outputFilePath, mergedBytes)
                End If
            End If
        End If

    End Sub

    Private Sub SimpleButton10_Click(sender As Object, e As EventArgs) Handles SimpleButton10.Click
        Dim imageFilePath As String = ""
        Dim pdfFilePath As String = ""
        Dim outputFilePath As String = ""

        OpenFileDialog1.Filter = "Image Files|*.jpg;*.jpeg|All Files|*.*"

        ' Open the image file dialog
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            imageFilePath = OpenFileDialog1.FileName

            OpenFileDialog1.Filter = "Excel Files|*.xlsx|All Files|*.*"

            ' Open the PDF file dialog
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                pdfFilePath = OpenFileDialog1.FileName

                SaveFileDialog1.Filter = "JPEG Files|*.jpg|All Files|*.*"

                ' Check if the output file already exists
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    outputFilePath = Path.ChangeExtension(SaveFileDialog1.FileName, ".jpg")

                    ' Check if the output file already exists
                    If File.Exists(outputFilePath) Then
                        Dim result As DialogResult = MessageBox.Show("The output file already exists. Do you want to replace it?", "File already exists", MessageBoxButtons.YesNo)
                        If result = DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    ' Read the image and PDF files as byte arrays
                    Dim imageBytes As Byte() = File.ReadAllBytes(imageFilePath)
                    Dim pdfBytes As Byte() = File.ReadAllBytes(pdfFilePath)

                    ' Find the end of the image in the image byte array
                    Dim imageEndIndex As Integer = -1
                    Dim imageEndSequence As Byte() = {&HFF, &HD9}
                    For i As Integer = imageBytes.Length - imageEndSequence.Length To 0 Step -1
                        Dim found As Boolean = True
                        For j As Integer = 0 To imageEndSequence.Length - 1
                            If imageBytes(i + j) <> imageEndSequence(j) Then
                                found = False
                                Exit For
                            End If
                        Next
                        If found Then
                            imageEndIndex = i + imageEndSequence.Length
                            Exit For
                        End If
                    Next

                    If imageEndIndex = -1 Then
                        MsgBox("End of image not found")
                        Exit Sub
                    End If

                    ' Create a new byte array to hold the merged image and PDF
                    Dim mergedBytes(imageEndIndex + pdfBytes.Length - 1) As Byte

                    ' Copy the image bytes to the merged byte array
                    Array.Copy(imageBytes, mergedBytes, imageEndIndex)

                    ' Copy the PDF bytes to the merged byte array
                    Array.Copy(pdfBytes, 0, mergedBytes, imageEndIndex, pdfBytes.Length)

                    ' Write the merged byte array to a new image file
                    File.WriteAllBytes(outputFilePath, mergedBytes)
                End If
            End If
        End If
    End Sub

    Private Sub SimpleButton9_Click(sender As Object, e As EventArgs) Handles SimpleButton9.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|excel Files (*.xlsx)|*.xlsx"
        If openFileDialog1.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        Dim inputFilePath As String = openFileDialog1.FileName

        Dim inputFileBytes As Byte() = File.ReadAllBytes(inputFilePath)

        Dim pdfStartSequence As Byte() = {&H50, &H4B, &H3, &H4}

        Dim pdfStartIndex As Integer = -1
        For i As Integer = 0 To inputFileBytes.Length - pdfStartSequence.Length
            Dim found As Boolean = True
            For j As Integer = 0 To pdfStartSequence.Length - 1
                If inputFileBytes(i + j) <> pdfStartSequence(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then
                pdfStartIndex = i
                Exit For
            End If
        Next

        If pdfStartIndex = -1 Then
            MsgBox("Excle not found in image")
            Exit Sub
        End If

        Dim pdfBytes(inputFileBytes.Length - pdfStartIndex - 1) As Byte
        Array.Copy(inputFileBytes, pdfStartIndex, pdfBytes, 0, pdfBytes.Length)

        Dim SaveFileDialog2 As New SaveFileDialog()
        SaveFileDialog2.Filter = "Excle Files (*.xlsx)|*.xlsx"

        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog2.FileName
            File.WriteAllBytes(outputFilePath, pdfBytes)
        End If

        Dim imageBytes(inputFileBytes.Length - pdfBytes.Length - 1) As Byte
        Array.Copy(inputFileBytes, 0, imageBytes, 0, pdfStartIndex)

        Dim SaveFileDialog3 As New SaveFileDialog()
        SaveFileDialog3.Filter = "JPEG Files (*.jpg)|*.jpg"

        If SaveFileDialog3.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog3.FileName
            File.WriteAllBytes(outputFilePath, imageBytes)
        End If
    End Sub


End Class
