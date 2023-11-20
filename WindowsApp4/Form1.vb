Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class Form1

    Private Sub EncryptFile(ByVal inputFile As String, ByVal outputFile As String, ByVal password As String)
        ' Create the encryption key from the password
        Dim key As Byte() = Encoding.UTF8.GetBytes(password)

        ' Initialize the AES encryption algorithm with the key
        Dim aesAlg As Aes = Aes.Create()
        aesAlg.Key = key

        ' Generate a random initialization vector (IV)
        aesAlg.GenerateIV()

        ' Write the IV to the output file
        Using outputStream As FileStream = New FileStream(outputFile, FileMode.Create)
            outputStream.Write(aesAlg.IV, 0, aesAlg.IV.Length)
        End Using

        ' Encrypt the input file using AES
        Using inputStream As FileStream = New FileStream(inputFile, FileMode.Open)
            Using cryptoStream As CryptoStream = New CryptoStream(File.Create(outputFile), aesAlg.CreateEncryptor(), CryptoStreamMode.Write)
                inputStream.CopyTo(cryptoStream)
            End Using
        End Using
    End Sub

    Private Sub DecryptFile(ByVal inputFile As String, ByVal outputFile As String, ByVal password As String)
        ' Create the encryption key from the password
        Dim key As Byte() = Encoding.UTF8.GetBytes(password)

        ' Initialize the AES encryption algorithm with the key and IV from the input file
        Dim aesAlg As Aes = Aes.Create()
        Using inputStream As FileStream = New FileStream(inputFile, FileMode.Open)
            Dim iv As Byte() = New Byte(aesAlg.IV.Length - 1) {}
            inputStream.Read(iv, 0, aesAlg.IV.Length)
            aesAlg.IV = iv

            ' Decrypt the input file using AES
            Using cryptoStream As CryptoStream = New CryptoStream(File.Create(outputFile), aesAlg.CreateDecryptor(key, aesAlg.IV), CryptoStreamMode.Write)
                inputStream.CopyTo(cryptoStream)
            End Using
        End Using
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

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

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' فتح ملف الصورة وملف ال PDF
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|PDF Files (*.pdf)|*.pdf"
        If openFileDialog1.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If
        ' تحديد مسار الملف الذي يحتوي على الصورة والملف الثاني (PDF)
        Dim inputFilePath As String = openFileDialog1.FileName
        ' قراءة بايتات الملف
        Dim inputFileBytes As Byte() = File.ReadAllBytes(inputFilePath)
        ' البايتات التي تشير إلى بداية الملف الثاني (PDF)
        Dim pdfStartSequence As Byte() = {&H25, &H50, &H44, &H46}
        ' البحث عن بداية الملف الثاني (PDF)
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
        ' التحقق من وجود بداية الملف الثاني (PDF)
        If pdfStartIndex = -1 Then
            Console.WriteLine("PDF not found in image")
            Exit Sub
        End If
        ' استخراج بايتات الملف الثاني (PDF)
        Dim pdfBytes(inputFileBytes.Length - pdfStartIndex - 1) As Byte
        Array.Copy(inputFileBytes, pdfStartIndex, pdfBytes, 0, pdfBytes.Length)
        ' حفظ الملف الثاني (PDF)
        Dim SaveFileDialog2 As New SaveFileDialog()
        SaveFileDialog2.Filter = "PDF Files (*.pdf)|*.pdf"
        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog2.FileName
            File.WriteAllBytes(outputFilePath, pdfBytes)
        End If
        ' استخراج بايتات الصورة بدون الملف الثاني (PDF)
        Dim imageBytes(inputFileBytes.Length - pdfBytes.Length - 1) As Byte
        Array.Copy(inputFileBytes, 0, imageBytes, 0, pdfStartIndex)
        ' حفظ الصورة بدون الملف الثاني (PDF)
        Dim SaveFileDialog3 As New SaveFileDialog()
        SaveFileDialog3.Filter = "JPEG Files (*.jpg)|*.jpg"
        If SaveFileDialog3.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog3.FileName
            File.WriteAllBytes(outputFilePath, imageBytes)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Dim image1Bytes As Byte() = File.ReadAllBytes("C:\Users\alicr\OneDrive\Desktop\New folder (2)\JPG\1.jpg")
        Dim image2Bytes As Byte() = File.ReadAllBytes("C:\Users\alicr\OneDrive\Desktop\New folder (2)\JPG\2.jpg")

        Dim startMarker As Byte() = New Byte() {&HFF, &HD8, &HFF, &HE0}
        Dim endMarker As Byte() = New Byte() {&HFF, &HD9}

        Dim startOffset As Integer = -1
        Dim endOffset As Integer = -1

        ' البحث عن علامات البداية والنهاية في الصورة الأولى
        For i As Integer = 0 To image1Bytes.Length - startMarker.Length - 1
            Dim match As Boolean = True
            For j As Integer = 0 To startMarker.Length - 1
                If image1Bytes(i + j) <> startMarker(j) Then
                    match = False
                    Exit For
                End If
            Next
            If match Then
                startOffset = i
                Exit For
            End If
        Next

        For i As Integer = image1Bytes.Length - endMarker.Length To 0 Step -1
            Dim match As Boolean = True
            For j As Integer = 0 To endMarker.Length - 1
                If image1Bytes(i + j) <> endMarker(j) Then
                    match = False
                    Exit For
                End If
            Next
            If match Then
                endOffset = i + endMarker.Length - 1
                Exit For
            End If
        Next

        ' التأكد من أن العثور على العلامات ناجح
        If startOffset = -1 Then
            Throw New Exception("Could not find start marker in image1")
        End If

        If endOffset = -1 Then
            Throw New Exception("Could not find end marker in image1")
        End If

        ' إنشاء MemoryStream لحفظ الصورة المدمجة
        Dim mergedStream As New MemoryStream()

        ' كتابة الجزء الأول من الصورة الأولى إلى mergedStream
        mergedStream.Write(image1Bytes, 0, startOffset + startMarker.Length)

        ' كتابة الجزء المشترك من الصورتين إلى mergedStream
        Dim commonBytesCount As Integer = image2Bytes.Length - endMarker.Length
        mergedStream.Write(image2Bytes, startMarker.Length, commonBytesCount)

        ' كتابة الجزء الأخير من الصورة الأولى إلى mergedStream
        mergedStream.Write(image1Bytes, endOffset, image1Bytes.Length - endOffset)

        ' حفظ الصورة المدمجة إلى ملف
        File.WriteAllBytes("C:\Users\alicr\OneDrive\Desktop\New folder (2)\JPG\333.jpg", mergedStream.ToArray())




    End Sub
    Sub ADD()
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
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ADD()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|word Files (*.docx)|*.docx"
        If openFileDialog1.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        ' تحديد مسار الملف الذي يحتوي على الصورة والملف الثاني (PDF)
        Dim inputFilePath As String = openFileDialog1.FileName

        ' قراءة بايتات الملف
        Dim inputFileBytes As Byte() = File.ReadAllBytes(inputFilePath)

        ' البايتات التي تشير إلى بداية الملف الثاني (PDF)
        Dim pdfStartSequence As Byte() = {&H50, &H4B, &H3, &H4}

        ' البحث عن بداية الملف الثاني (PDF)
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

        ' التحقق من وجود بداية الملف الثاني (PDF)
        If pdfStartIndex = -1 Then
            MsgBox("docx not found in image")
            Exit Sub
        End If

        ' استخراج بايتات الملف الثاني (PDF)
        Dim pdfBytes(inputFileBytes.Length - pdfStartIndex - 1) As Byte
        Array.Copy(inputFileBytes, pdfStartIndex, pdfBytes, 0, pdfBytes.Length)

        ' حفظ الملف الثاني (PDF)
        Dim SaveFileDialog2 As New SaveFileDialog()
        SaveFileDialog2.Filter = "docx Files (*.docx)|*.docx"

        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog2.FileName
            File.WriteAllBytes(outputFilePath, pdfBytes)
        End If

        ' استخراج بايتات الصورة بدون الملف الثاني (PDF)
        Dim imageBytes(inputFileBytes.Length - pdfBytes.Length - 1) As Byte
        Array.Copy(inputFileBytes, 0, imageBytes, 0, pdfStartIndex)

        ' حفظ الصورة بدون الملف الثاني (PDF)
        Dim SaveFileDialog3 As New SaveFileDialog()
        SaveFileDialog3.Filter = "JPEG Files (*.jpg)|*.jpg"

        If SaveFileDialog3.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog3.FileName
            File.WriteAllBytes(outputFilePath, imageBytes)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ADD()
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        OpenFileDialog1.Filter = "JPEG Files (*.jpg)|*.jpg|RAR Files (*.RAR)|*.RAR"
        If OpenFileDialog1.ShowDialog() <> DialogResult.OK Then
            Exit Sub
        End If

        ' تحديد مسار الملف الذي يحتوي على الصورة والملف الثاني (PDF)
        Dim inputFilePath As String = OpenFileDialog1.FileName

        ' قراءة بايتات الملف
        Dim inputFileBytes As Byte() = File.ReadAllBytes(inputFilePath)

        ' البايتات التي تشير إلى بداية الملف الثاني (PDF)
        Dim pdfStartSequence As Byte() = {&H52, &H61, &H72, &H21}

        ' البحث عن بداية الملف الثاني (PDF)
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

        ' التحقق من وجود بداية الملف الثاني (PDF)
        If pdfStartIndex = -1 Then
            MsgBox("RAR not found in image")
            Exit Sub
        End If

        ' استخراج بايتات الملف الثاني (PDF)
        Dim pdfBytes(inputFileBytes.Length - pdfStartIndex - 1) As Byte
        Array.Copy(inputFileBytes, pdfStartIndex, pdfBytes, 0, pdfBytes.Length)

        ' حفظ الملف الثاني (PDF)
        Dim SaveFileDialog2 As New SaveFileDialog()
        SaveFileDialog2.Filter = "RAR Files (*.RAR)|*.RAR"

        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog2.FileName
            File.WriteAllBytes(outputFilePath, pdfBytes)
        End If

        ' استخراج بايتات الصورة بدون الملف الثاني (PDF)
        Dim imageBytes(inputFileBytes.Length - pdfBytes.Length - 1) As Byte
        Array.Copy(inputFileBytes, 0, imageBytes, 0, pdfStartIndex)

        ' حفظ الصورة بدون الملف الثاني (PDF)
        Dim SaveFileDialog3 As New SaveFileDialog()
        SaveFileDialog3.Filter = "JPEG Files (*.jpg)|*.jpg"

        If SaveFileDialog3.ShowDialog() = DialogResult.OK Then
            Dim outputFilePath As String = SaveFileDialog3.FileName
            File.WriteAllBytes(outputFilePath, imageBytes)
        End If
    End Sub
End Class
