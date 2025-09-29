Imports System.Text.RegularExpressions
Imports System.IO
Public Class Form1
    Dim LangSel As Boolean
    Dim FileSel As Boolean
    'Dim IsNetwork As Boolean
    Dim langRex As String
    Dim OnlyLang As String
    Dim FilePath As String
    Dim IsPrep As Boolean
    Dim OriginalFileNameExt As String
    Dim OriginalFileNameNoExt As String
    Dim OriginalPathNoName As String
    Dim NoLatin As Boolean
    'GGGGGG

    Private Sub ListBox1_DragDrop(sender As Object, e As DragEventArgs) Handles ListBox1.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        Dim sep() As Char = {"/", "\", "//"}

        For Each path In files
            If ListBox1.Items.Count > 0 Then
                Exit Sub
            End If

            ' NEW
            'Determine if is a prep or post
            'MsgBox(Strings.Left(Strings.Right(path, 15), 9))
            If Strings.Left(Strings.Right(path, 18), 9) = "_Totrans_" Then
                IsPrep = False
                Button1.Text = "Post Process"
                Button1.ForeColor = Color.Green
                Label1.Text = "Drop xml file to panel and click Post Process button"
                Label1.BackColor = Color.LightSeaGreen
                'MsgBox("Is a post")
            Else
                IsPrep = True
                Button1.Text = "Prepare File"
                'MsgBox("Is a prep")
            End If
            If Strings.Right(Strings.LCase(path), 4) = ".xml" Then
                ListBox1.Items.Add(path)
                FilePath = path
                'If Strings.Left(FilePath, 2) = "\\" Then
                'IsNetwork = True
                'End If
                OriginalFileNameExt = path.Split(sep).Last()
                OriginalFileNameNoExt = Strings.Left(OriginalFileNameExt, Strings.Len(OriginalFileNameExt) - 4)
                OriginalPathNoName = Strings.Left(FilePath, Strings.Len(FilePath) - Strings.Len(OriginalFileNameExt))
                FileSel = True
                If IsPrep And LangSel = True And FileSel = True Then
                    Button1.Enabled = True
                    ComboBox1.Enabled = True
                ElseIf IsPrep = False And FileSel = True Then
                    Button1.Enabled = True
                    ComboBox1.Enabled = False
                End If
            End If
        Next
    End Sub

    Private Sub ListBox1_DragEnter(sender As Object, e As DragEventArgs) Handles ListBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim pattern As String
        Dim replacement As String

        'Reads the whole file
        Dim value As String = File.ReadAllText(FilePath)

        'Checks if WIPO XML else exit
        If value.Contains("ST26SequenceListing_V1_3.dtd") = False Then
            MsgBox("Error: This XML seems not to be a WIPO XML or version is incorrect." & vbCrLf & "Ask to Localization dept.", MessageBoxIcon.Error)
            'Application.Exit()
            'Stop
            Sal()

        End If

        'If is a Prep
        If IsPrep = True Then
            'Special Case when client enters Aplicant name as DE
            pattern = "(<ApplicantName languageCode=)(""de"">)(.+?)(</ApplicantName>)(\r\n)"
            replacement = "$1""en"">$3$4$5"
            Dim rgx1 As New Regex(pattern)
            value = rgx1.Replace(value, replacement)

            'Check if some language code is NOT english (en) if different warning and exit
            Dim m As Match = Regex.Match(value, "languageCode="".+""", RegexOptions.IgnoreCase)
            While m.Success
                Dim strLanCode As String = m.Value
                Console.WriteLine(strLanCode)
                strLanCode = strLanCode.Substring(14, 2)
                If strLanCode.ToLower <> "en" Then
                    MsgBox("Error: This WIPO XML has a non English lang code=" & strLanCode & vbCrLf & "Ask to Localization dept.", MessageBoxIcon.Exclamation)
                    'Application.Exit()
                    'Stop
                    Sal()
                End If
                'MsgBox(strLanCode)
                m = m.NextMatch
            End While

            'APPLICANT NAME (only if language is zh)

            'Special case where client send a file with ApplicantNameLatin included we change lang en to final lang only
            If value.Contains("<ApplicantNameLatin>") = True And NoLatin = True Then
                pattern = "(<ApplicantName languageCode=)(""en"">)(.+?)(</ApplicantName>)(\r\n\t+)"
                replacement = "$1" & langRex & ">$3$4$5"
                Dim rgx As New Regex(pattern)
                value = rgx.Replace(value, replacement)
            End If

            'Now when ApplicantNameLatin does not exist (the usual)
            If value.Contains("<ApplicantNameLatin>") = False And NoLatin = True Then
                pattern = "(<ApplicantName languageCode=)(""en"">)(.+?)(</ApplicantName>)(\r\n\t+)"
                replacement = "$1" & langRex & ">$3$4$5<ApplicantNameLatin>$3</ApplicantNameLatin>$5"
                Dim rgx As New Regex(pattern)
                value = rgx.Replace(value, replacement)
            End If

            'INVENTOR NAME (only if language is zh)

            'Special case where client send a file with InvertorNameLatin included we change lang en to final lang only
            If value.Contains("<InventorNameLatin") = True And NoLatin = True Then
                pattern = "(<InventorName languageCode=)(""en"">)(.+?)(</InventorName>)(\r\n\t+)"
                replacement = "$1" & langRex & ">$3$4$5"
                Dim rgx5 As New Regex(pattern)
                value = rgx5.Replace(value, replacement)
            End If

            'Now when InventorNameLatin does not exist (the usual)
            If value.Contains("<InventorNameLatin") = False And NoLatin = True Then
                pattern = "(<InventorName languageCode=)(""en"">)(.+?)(</InventorName>)(\r\n\t+)"
                replacement = "$1" & langRex & ">$3$4$5<InventorNameLatin>$3</InventorNameLatin>$5"
                Dim rgx5 As New Regex(pattern)
                value = rgx5.Replace(value, replacement)
            End If

            'Invention title
            pattern = "(<InventionTitle languageCode=)(""en"">)(.+?)(</InventionTitle>)(\r\n\t+)"
            replacement = "$1$2$3$4$5$1" & langRex & ">$3$4$5"
            Dim rgx2 As New Regex(pattern)
            value = rgx2.Replace(value, replacement)

            'Header
            pattern = "(<ST26SequenceListing)(.+?)(dtdVersion.+?>)"
            replacement = "$1 originalFreeTextLanguageCode=""en"" nonEnglishFreeTextLanguageCode=" & langRex & " $3"
            Dim rgx3 As New Regex(pattern)
            value = rgx3.Replace(value, replacement)

            'Sequences: if lang is zh-tw no sequences will be translated
            If OnlyLang <> "zh-tw" Then
                pattern = "(<INSDQualifier_name>)(bound_moiety|cell_type|clone|clone_lib|collected_by|cultivar|dev_stage|ecotype|function|gene_synonym|haplogroup|host|identified_by|isolate|isolation_source|lab_host|mating_type|note|organism|phenotype|pop_variant|product|serotype|serovar|sex|standard_name|strain|sub_clone|sub_species|sub_strain|tissue_lib|tissue_type|variety)(</INSDQualifier_name>)(\r\n)(\t+)(<INSDQualifier_value>)(.+?)(</INSDQualifier_value>)"
                replacement = "$1$2$3$4$5$6$7$8$4$5<NonEnglishQualifier_value>$7</NonEnglishQualifier_value>"
                Dim rgx4 As New Regex(pattern)
                value = rgx4.Replace(value, replacement)
                'System.Diagnostics.Debug.WriteLine(value)
            End If

            'Writes Final file
            File.WriteAllText(OriginalPathNoName & OriginalFileNameNoExt & "_Totrans_" & OnlyLang & ".xml", value)

            'Write Log if zh-tw warms user about not generate qualifiers.
            Dim Logitem As String
            Logitem = ComboBox1.SelectedItem & "............."
            If OnlyLang <> "zh-tw" Then
                Logitem = Strings.Left(Logitem, 45) & " Done....Please keep the generated file name in all localization process. Do not rename."
            Else
                Logitem = Strings.Left(Logitem, 45) & " Done....Please keep the generated file name in all localization process. Do not rename. Qualifiers not generated for zh-tw."
            End If
            ListBox2.Items.Add(Logitem)

        End If

        If IsPrep = False Then
            Dim NewFileNameNoExt As String

            'UTF FIX
            pattern = "encoding=""utf-8"
            replacement = "encoding=""UTF-8"
            Dim rgx5 As New Regex(pattern)
            value = rgx5.Replace(value, replacement)
            NewFileNameNoExt = OriginalFileNameNoExt.Replace("Totrans", "FINAL")

            'CHARACTER FIX
            ' Patrón de expresión regular para encontrar las cadenas entre Tags que contienen ciertos caracteres llama a funcion ReplaceChars
            pattern = "<NonEnglishQualifier_value>(.+?)</NonEnglishQualifier_value>"
            Dim rgx6 As New Regex(pattern)
            value = rgx6.Replace(value, pattern, Function(match) ReplaceChars(match.Value))



            'Writes Final file
            File.WriteAllText(OriginalPathNoName & NewFileNameNoExt & ".xml", value)

            'Write Log
            Dim Logitem As String
            Logitem = ComboBox1.SelectedItem & "Final File: " & NewFileNameNoExt & ".xml" & " ready to be delivered. Rename as original if you need."
            ListBox2.Items.Add(Logitem)

            'Limpia la lista
            ListBox1.Items.Clear()
            Button1.Enabled = False
        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim Lang As String
        NoLatin = False

        Lang = ComboBox1.SelectedItem
        langRex = """" & Strings.Left(Strings.Right(ComboBox1.SelectedItem, 5), 2) & """"
        OnlyLang = Strings.Right(ComboBox1.SelectedItem, 5)

        'Here Nolatin determines if language needs translation of Applicante name and Inventor name
        'MessageBox.Show(OnlyLang)
        If OnlyLang = "zh-cn" Or OnlyLang = "zh-tw" Then
            NoLatin = True
        Else
            NoLatin = False
        End If



        LangSel = True
        If LangSel = True And FileSel = True Then
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.TopMost = True
        Versionlbl.Text = "Version: " & Application.ProductVersion
    End Sub

    Public Sub Sal()
        Close()
        Application.Exit()
        Stop
        End
    End Sub
    Function ReplaceChars(input As String) As String
        'Reemplaza “(&H201C) y ”(&H201D) por &quot; y ’(&H2019) por &apos;
        Return input.Replace(ChrW(&H201C), "&quot;").Replace(ChrW(&H201D), "&quot;").Replace(ChrW(&H2019), "&apos;")
    End Function
End Class
