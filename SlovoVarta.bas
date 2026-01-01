'===============================================================================
' Slovo-Varta (Слово-Варта)
' Version:  1.0.0a
' Date: 2026-01-01
'
' Description:
'   Open-source VBA module for Microsoft Excel designed for the automatic
'   declension of Ukrainian names, surnames, and patronymics.
'
' Purpose:
'   Administrative work in the Ukrainian military and public sector often
'   requires processing thousands of names in different grammatical cases.
'   This project automates that process, reducing manual errors and saving
'   time for more critical tasks.
'
' Author:
'   Yuriy Novostavskiy (@yurnov)
'   Created with support of GitHub Copilot (Claude Sonnet 4.5 model)
'
' License: MIT License
' Repository: https://github.com/yurnov/slovo-varta
'
' Inspired by: shevchenko-js library (https://github.com/tooleks/shevchenko-js)
'
' Special Thanks:
'   - Authors of shevchenko-js library
'   - Defense Forces of Ukraine (Сили оборони України)
'
' Support:
'   If you find this project helpful, please consider donating to:
'   "Come Back Alive" Foundation (Повернись живим)
'   https://savelife.in.ua/en/donate-en
'   Comprehensively equips the Defence Forces of Ukraine with equipment
'   and implements educational projects for the military.
'
'===============================================================================
' MIT License
'
' Copyright (c) 2025 Slovo-Varta Contributors
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'===============================================================================

Option Explicit

'===============================================================================
' ENUMERATIONS
'===============================================================================

Public Enum GrammaticalCase
    Nominative = 1  ' Називний відмінок (хто?  що?)
    Genitive = 2    ' Родовий відмінок (кого? чого?)
    Dative = 3      ' Давальний відмінок (кому? чому?)
End Enum

Public Enum GrammaticalGender
    Masculine = 1   ' Чоловічий рід
    Feminine = 2    ' Жіночий рід
End Enum

Public Enum NameComponent
    GivenName = 1       ' Ім'я
    PatronymicName = 2  ' По батькові
    FamilyName = 3      ' Прізвище
End Enum

'===============================================================================
' PUBLIC API - MAIN FUNCTIONS
'===============================================================================

'/**
' * Universal declension function
' *
' * @param nameText - Name component in nominative case
' * @param nameType - Type:  "given"/"patronymic"/"family"
' * @param gender - Gender: "m"/"f"/"ч"/"ж"/"masculine"/"feminine"
' * @param targetCase - Target case: "genitive"/"dative"
' * @return Declined name in specified case
' *
' * Example:
' *   =DeclineName("Шевченко", "family", "m", "dative")
' *   Returns: "Шевченку"
' */
Public Function DeclineName(ByVal nameText As String, _
                           ByVal nameType As String, _
                           ByVal gender As String, _
                           ByVal targetCase As String) As String
    On Error GoTo ErrorHandler

    nameText = Trim(nameText)
    nameType = Trim(nameType)
    gender = Trim(gender)
    targetCase = Trim(targetCase)

    If nameText = "" Then
        DeclineName = ""
        Exit Function
    End If

    Dim genderEnum As GrammaticalGender
    genderEnum = ParseGender(gender)

    Dim caseEnum As GrammaticalCase
    caseEnum = ParseCase(targetCase)

    Dim nameTypeEnum As NameComponent
    nameTypeEnum = ParseNameType(nameType)

    DeclineName = DeclineNameInternal(nameText, nameTypeEnum, genderEnum, caseEnum)
    Exit Function

ErrorHandler:
    DeclineName = "#ERROR:  " & Err.Description
End Function

'/**
' * Decline given name (ім'я) to genitive case
' *
' * @param givenName - Given name in nominative case
' * @param gender - Gender (m/f/ч/ж)
' * @return Given name in genitive case (родовий відмінок)
' *
' * Example:
' *   =GivenNameGenitive("Тарас", "m")
' *   Returns: "Тараса"
' */
Public Function GivenNameGenitive(ByVal givenName As String, ByVal gender As String) As String
    GivenNameGenitive = DeclineName(givenName, "given", gender, "genitive")
End Function

'/**
' * Decline given name (ім'я) to dative case
' *
' * @param givenName - Given name in nominative case
' * @param gender - Gender (m/f/ч/ж)
' * @return Given name in dative case (давальний відмінок)
' *
' * Example:
' *   =GivenNameDative("Тарас", "m")
' *   Returns: "Тарасу"
' */
Public Function GivenNameDative(ByVal givenName As String, ByVal gender As String) As String
    GivenNameDative = DeclineName(givenName, "given", gender, "dative")
End Function

'/**
' * Decline patronymic (по батькові) to genitive case
' *
' * Example:
' *   =PatronymicGenitive("Григорович", "m")
' *   Returns: "Григоровича"
' */
Public Function PatronymicGenitive(ByVal patronymic As String, ByVal gender As String) As String
    PatronymicGenitive = DeclineName(patronymic, "patronymic", gender, "genitive")
End Function

'/**
' * Decline patronymic (по батькові) to dative case
' *
' * Example:
' *   =PatronymicDative("Григорович", "m")
' *   Returns: "Григоровичу"
' */
Public Function PatronymicDative(ByVal patronymic As String, ByVal gender As String) As String
    PatronymicDative = DeclineName(patronymic, "patronymic", gender, "dative")
End Function

'/**
' * Decline family name (прізвище) to genitive case
' *
' * Example:
' *   =FamilyNameGenitive("Шевченко", "m")
' *   Returns: "Шевченка"
' */
Public Function FamilyNameGenitive(ByVal familyName As String, ByVal gender As String) As String
    FamilyNameGenitive = DeclineName(familyName, "family", gender, "genitive")
End Function

'/**
' * Decline family name (прізвище) to dative case
' *
' * Example:
' *   =FamilyNameDative("Шевченко", "m")
' *   Returns: "Шевченку"
' */
Public Function FamilyNameDative(ByVal familyName As String, ByVal gender As String) As String
    FamilyNameDative = DeclineName(familyName, "family", gender, "dative")
End Function

'===============================================================================
' DEBUG AND UTILITY FUNCTIONS
'===============================================================================

'/**
' * Debug function to diagnose declension issues
' *
' * Returns detailed information about the declension process
' */
Public Function DebugDecline(ByVal nameText As String, _
                            ByVal nameType As String, _
                            ByVal gender As String, _
                            ByVal targetCase As String) As String
    Dim result As String
    Dim nameLen As Long

    result = "=== SLOVO-VARTA DEBUG ===" & vbCrLf
    result = result & "Raw Input: [" & nameText & "]" & vbCrLf
    result = result & "Trimmed:  [" & Trim(nameText) & "]" & vbCrLf
    result = result & "Length: " & Len(Trim(nameText)) & vbCrLf
    result = result & "Name Type: [" & nameType & "]" & vbCrLf
    result = result & "Gender:  [" & gender & "]" & vbCrLf
    result = result & "Case: [" & targetCase & "]" & vbCrLf

    nameLen = Len(Trim(nameText))
    If nameLen > 0 Then
        result = result & "Last char: [" & Right(Trim(nameText), 1) & "]" & vbCrLf
        result = result & "Last 2 chars: [" & Right(Trim(nameText), 2) & "]" & vbCrLf
        result = result & "Last 3 chars: [" & Right(Trim(nameText), 3) & "]" & vbCrLf
    End If

    On Error Resume Next
    result = result & "Result: [" & DeclineName(nameText, nameType, gender, targetCase) & "]" & vbCrLf
    If Err.Number <> 0 Then
        result = result & "ERROR: " & Err.Description & vbCrLf
    End If
    On Error GoTo 0

    DebugDecline = result
End Function

'/**
' * Get version information
' */
Public Function SlovoVartaVersion() As String
    SlovoVartaVersion = "Slovo-Varta v1.0.0a - Ukrainian Name Declension for Excel"
End Function

'===============================================================================
' INTERNAL PARSING FUNCTIONS
'===============================================================================

Private Function ParseGender(ByVal genderStr As String) As GrammaticalGender
    Dim g As String
    g = LCase(Trim(genderStr))

    Select Case g
        Case "m", "masculine", "man", "male"
            ParseGender = Masculine
        Case "f", "feminine", "woman", "w", "female"
            ParseGender = Feminine
        Case Else
            ' Check Ukrainian characters using ChrW codes for proper Unicode support
            If g = ChrW(1095) Or _
               g = ChrW(1095) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1074) & ChrW(1110) & ChrW(1082) Or _
               g = ChrW(1095) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1074) & ChrW(1110) & ChrW(1095) & ChrW(1080) & ChrW(1081) Then
                ' ч, чоловік, чоловічий
                ParseGender = Masculine
            ElseIf g = ChrW(1078) Or _
                   g = ChrW(1078) & ChrW(1110) & ChrW(1085) & ChrW(1082) & ChrW(1072) Or _
                   g = ChrW(1078) & ChrW(1110) & ChrW(1085) & ChrW(1086) & ChrW(1095) & ChrW(1080) & ChrW(1081) Then
                ' ж, жінка, жіночий
                ParseGender = Feminine
            Else
                Err.Raise vbObjectError + 1, "ParseGender", _
                    "Invalid gender: '" & genderStr & "'.  Use:  m/f/male/female/ч/ж/masculine/feminine"
            End If
    End Select
End Function

Private Function ParseCase(ByVal caseStr As String) As GrammaticalCase
    Dim c As String
    c = LCase(Trim(caseStr))

    Select Case c
        Case "genitive", "gen", "g"
            ParseCase = Genitive
        Case "dative", "dat", "d"
            ParseCase = Dative
        Case "nominative", "nom", "n"
            ParseCase = Nominative
        Case Else
            Err.Raise vbObjectError + 2, "ParseCase", _
                "Invalid case: '" & caseStr & "'. Use: genitive/dative"
    End Select
End Function

Private Function ParseNameType(ByVal nameTypeStr As String) As NameComponent
    Dim nt As String
    nt = LCase(Trim(nameTypeStr))

    Select Case nt
        Case "given", "givenname", "first", "firstname", "g"
            ParseNameType = GivenName
        Case "patronymic", "patronymicname", "middle", "p"
            ParseNameType = PatronymicName
        Case "family", "familyname", "last", "lastname", "surname", "f"
            ParseNameType = FamilyName
        Case Else
            Err.Raise vbObjectError + 3, "ParseNameType", _
                "Invalid name type: '" & nameTypeStr & "'. Use: given/patronymic/family"
    End Select
End Function

'===============================================================================
' CORE DECLENSION LOGIC
'===============================================================================

Private Function DeclineNameInternal(ByVal nameText As String, _
                                    ByVal nameType As NameComponent, _
                                    ByVal gender As GrammaticalGender, _
                                    ByVal targetCase As GrammaticalCase) As String
    If targetCase = Nominative Then
        DeclineNameInternal = nameText
        Exit Function
    End If

    If InStr(nameText, "-") > 0 Then
        DeclineNameInternal = DeclineCompoundName(nameText, nameType, gender, targetCase)
        Exit Function
    End If

    Select Case nameType
        Case GivenName
            DeclineNameInternal = DeclineGivenName(nameText, gender, targetCase)
        Case PatronymicName
            DeclineNameInternal = DeclinePatronymic(nameText, gender, targetCase)
        Case FamilyName
            DeclineNameInternal = DeclineFamilyName(nameText, gender, targetCase)
    End Select
End Function

'===============================================================================
' GIVEN NAME DECLENSION (Ім'я)
'===============================================================================

Private Function DeclineGivenName(ByVal name As String, _
                                 ByVal gender As GrammaticalGender, _
                                 ByVal targetCase As GrammaticalCase) As String
    Dim stem As String
    Dim nameLen As Long
    nameLen = Len(name)

    If gender = Masculine Then
        ' === -ій (Юрій, Андрій, Василій, Сергій) ===
        ' Keep 'і' + add ending (Юрій → Юрі + ю = Юрію)
        If EndsWith(name, ChrW(1110) & ChrW(1081)) Then ' ій
            stem = Left(name, nameLen - 1)  ' Remove only й, keep і
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1103) ' + я
            Else
                DeclineGivenName = stem & ChrW(1102) ' + ю
            End If
            Exit Function
        End If

        ' === -й (not preceded by і) ===
        If EndsWith(name, ChrW(1081)) And Not EndsWith(name, ChrW(1110) & ChrW(1081)) Then ' й
            stem = Left(name, nameLen - 1)
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1103) ' я
            Else
                DeclineGivenName = stem & ChrW(1102) ' ю
            End If
            Exit Function
        End If

        ' === -о (Павло, Данило) ===
        If EndsWith(name, ChrW(1086)) Then ' о
            stem = Left(name, nameLen - 1)
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1072) ' а
            Else
                DeclineGivenName = stem & ChrW(1091) ' у
            End If
            Exit Function
        End If

        ' === -ь (Ігор) ===
        If EndsWith(name, ChrW(1100)) Then ' ь
            stem = Left(name, nameLen - 1)
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1103) ' я
            Else
                DeclineGivenName = stem & ChrW(1102) ' ю
            End If
            Exit Function
        End If

        ' === -я, -а (rare masculine:  Ілля, Нікіта) ===
        If EndsWith(name, ChrW(1103)) Or EndsWith(name, ChrW(1072)) Then ' я or а
            stem = Left(name, nameLen - 1)
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1110) ' і
            Else
                DeclineGivenName = stem & ChrW(1110) ' і
            End If
            Exit Function
        End If

        ' === CONSONANT (Тарас, Іван, Богдан) ===
        If IsConsonant(Right(name, 1)) Then
            If targetCase = Genitive Then
                DeclineGivenName = name & ChrW(1072) ' а
            Else
                DeclineGivenName = name & ChrW(1091) ' у
            End If
            Exit Function
        End If

    Else  ' FEMININE
        ' === -а (Людмила, Лариса, Марія) ===
        If EndsWith(name, ChrW(1072)) Then ' а
            stem = Left(name, nameLen - 1)
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1080) ' и
            Else
                DeclineGivenName = stem & ChrW(1110) ' і
            End If
            Exit Function
        End If

        ' === -я (Софія, Наталія) ===
        If EndsWith(name, ChrW(1103)) Then ' я
            stem = Left(name, nameLen - 1)
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1111) ' ї
            Else
                DeclineGivenName = stem & ChrW(1111) ' ї
            End If
            Exit Function
        End If

        ' === -ь (Любов) ===
        If EndsWith(name, ChrW(1100)) Then ' ь
            stem = Left(name, nameLen - 1)
            If targetCase = Genitive Then
                DeclineGivenName = stem & ChrW(1110) ' і
            Else
                DeclineGivenName = stem & ChrW(1110) ' і
            End If
            Exit Function
        End If
    End If

    ' Default:  no change (indeclinable names)
    DeclineGivenName = name
End Function

'===============================================================================
' PATRONYMIC DECLENSION (По батькові)
'===============================================================================

Private Function DeclinePatronymic(ByVal patronymic As String, _
                                  ByVal gender As GrammaticalGender, _
                                  ByVal targetCase As GrammaticalCase) As String
    Dim stem As String
    Dim nameLen As Long
    nameLen = Len(patronymic)

    If gender = Masculine Then
        ' === Various -ович endings ===
        If EndsWith(patronymic, ChrW(1086) & ChrW(1074) & ChrW(1080) & ChrW(1095)) Or _
           EndsWith(patronymic, ChrW(1077) & ChrW(1074) & ChrW(1080) & ChrW(1095)) Or _
           EndsWith(patronymic, ChrW(1108) & ChrW(1074) & ChrW(1080) & ChrW(1095)) Or _
           EndsWith(patronymic, ChrW(1081) & ChrW(1086) & ChrW(1074) & ChrW(1080) & ChrW(1095)) Or _
           EndsWith(patronymic, ChrW(1100) & ChrW(1086) & ChrW(1074) & ChrW(1080) & ChrW(1095)) Or _
           EndsWith(patronymic, ChrW(1110) & ChrW(1086) & ChrW(1074) & ChrW(1080) & ChrW(1095)) Then
            ' ович, евич, євич, йович, ьович, іович
            stem = Left(patronymic, nameLen - 2)  ' Remove ич
            If targetCase = Genitive Then
                DeclinePatronymic = stem & ChrW(1080) & ChrW(1095) & ChrW(1072) ' ича
            Else
                DeclinePatronymic = stem & ChrW(1080) & ChrW(1095) & ChrW(1091) ' ичу
            End If
            Exit Function
        End If

        ' === -ич (standalone:  Ілліч) ===
        If EndsWith(patronymic, ChrW(1080) & ChrW(1095)) Then ' ич
            If targetCase = Genitive Then
                DeclinePatronymic = patronymic & ChrW(1072) ' а
            Else
                DeclinePatronymic = patronymic & ChrW(1091) ' у
            End If
            Exit Function
        End If

    Else  ' FEMININE
        ' === -івна, -ївна ===
        If EndsWith(patronymic, ChrW(1110) & ChrW(1074) & ChrW(1085) & ChrW(1072)) Or _
           EndsWith(patronymic, ChrW(1111) & ChrW(1074) & ChrW(1085) & ChrW(1072)) Then
            stem = Left(patronymic, nameLen - 1)  ' Remove а
            If targetCase = Genitive Then
                DeclinePatronymic = stem & ChrW(1080) ' и
            Else
                DeclinePatronymic = stem & ChrW(1110) ' і
            End If
            Exit Function
        End If

        ' === -ична (rare) ===
        If EndsWith(patronymic, ChrW(1080) & ChrW(1095) & ChrW(1085) & ChrW(1072)) Then
            stem = Left(patronymic, nameLen - 1)
            If targetCase = Genitive Then
                DeclinePatronymic = stem & ChrW(1080) ' и
            Else
                DeclinePatronymic = stem & ChrW(1110) ' і
            End If
            Exit Function
        End If
    End If

    ' Default: no change
    DeclinePatronymic = patronymic
End Function

'===============================================================================
' FAMILY NAME DECLENSION (Прізвище)
'===============================================================================

Private Function DeclineFamilyName(ByVal familyName As String, _
                                  ByVal gender As GrammaticalGender, _
                                  ByVal targetCase As GrammaticalCase) As String
    Dim stem As String
    Dim nameLen As Long
    nameLen = Len(familyName)

    If gender = Masculine Then
        ' === ADJECTIVES:  -ський, -цький ===
        If EndsWith(familyName, ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1080) & ChrW(1081)) Or _
           EndsWith(familyName, ChrW(1094) & ChrW(1100) & ChrW(1082) & ChrW(1080) & ChrW(1081)) Then
            stem = Left(familyName, nameLen - 2)  ' Remove ий
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1086) & ChrW(1075) & ChrW(1086) ' ого
            Else
                DeclineFamilyName = stem & ChrW(1086) & ChrW(1084) & ChrW(1091) ' ому
            End If
            Exit Function
        End If

        ' === ADJECTIVES: -ний, -ній, -ий, -ій ===
        If EndsWith(familyName, ChrW(1085) & ChrW(1080) & ChrW(1081)) Or _
           EndsWith(familyName, ChrW(1085) & ChrW(1110) & ChrW(1081)) Or _
           EndsWith(familyName, ChrW(1080) & ChrW(1081)) Or _
           EndsWith(familyName, ChrW(1110) & ChrW(1081)) Then
            stem = Left(familyName, nameLen - 2)
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1086) & ChrW(1075) & ChrW(1086) ' ого
            Else
                DeclineFamilyName = stem & ChrW(1086) & ChrW(1084) & ChrW(1091) ' ому
            End If
            Exit Function
        End If

        ' === -ко (Шевченко, Максименко) ===
        If EndsWith(familyName, ChrW(1082) & ChrW(1086)) Then
            stem = Left(familyName, nameLen - 1)
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1072) ' а
            Else
                DeclineFamilyName = stem & ChrW(1091) ' у
            End If
            Exit Function
        End If

        ' === -ук, -юк ===
        If EndsWith(familyName, ChrW(1091) & ChrW(1082)) Or _
           EndsWith(familyName, ChrW(1102) & ChrW(1082)) Then
            If targetCase = Genitive Then
                DeclineFamilyName = familyName & ChrW(1072) ' а
            Else
                DeclineFamilyName = familyName & ChrW(1091) ' у
            End If
            Exit Function
        End If

        ' === -ець ===
        If EndsWith(familyName, ChrW(1077) & ChrW(1094) & ChrW(1100)) Then
            stem = Left(familyName, nameLen - 3)
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1094) & ChrW(1103) ' ця
            Else
                DeclineFamilyName = stem & ChrW(1094) & ChrW(1102) ' цю
            End If
            Exit Function
        End If

        ' === -ич ===
        If EndsWith(familyName, ChrW(1080) & ChrW(1095)) Then
            If targetCase = Genitive Then
                DeclineFamilyName = familyName & ChrW(1072) ' а
            Else
                DeclineFamilyName = familyName & ChrW(1091) ' у
            End If
            Exit Function
        End If

        ' === CONSONANT (Коваль, Бондар) ===
        If IsConsonant(Right(familyName, 1)) Then
            If targetCase = Genitive Then
                DeclineFamilyName = familyName & ChrW(1072) ' а
            Else
                DeclineFamilyName = familyName & ChrW(1091) ' у
            End If
            Exit Function
        End If

    Else  ' FEMININE
        ' === ADJECTIVES: -ська, -цька ===
        If EndsWith(familyName, ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1072)) Or _
           EndsWith(familyName, ChrW(1094) & ChrW(1100) & ChrW(1082) & ChrW(1072)) Then
            stem = Left(familyName, nameLen - 1)  ' Remove а
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1086) & ChrW(1111) ' ої
            Else
                DeclineFamilyName = stem & ChrW(1110) & ChrW(1081) ' ій
            End If
            Exit Function
        End If

        ' === -ко ===
        If EndsWith(familyName, ChrW(1082) & ChrW(1086)) Then
            stem = Left(familyName, nameLen - 1)
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1072) ' а
            Else
                DeclineFamilyName = stem & ChrW(1091) ' у
            End If
            Exit Function
        End If

        ' === -а ===
        If EndsWith(familyName, ChrW(1072)) Then
            stem = Left(familyName, nameLen - 1)
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1080) ' и
            Else
                DeclineFamilyName = stem & ChrW(1110) ' і
            End If
            Exit Function
        End If

        ' === -я ===
        If EndsWith(familyName, ChrW(1103)) Then
            stem = Left(familyName, nameLen - 1)
            If targetCase = Genitive Then
                DeclineFamilyName = stem & ChrW(1111) ' ї
            Else
                DeclineFamilyName = stem & ChrW(1111) ' ї
            End If
            Exit Function
        End If

        ' === CONSONANT (indeclinable for women:  Косач) ===
        If IsConsonant(Right(familyName, 1)) Then
            DeclineFamilyName = familyName
            Exit Function
        End If
    End If

    ' Default: no change
    DeclineFamilyName = familyName
End Function

'===============================================================================
' COMPOUND NAMES (with hyphens)
'===============================================================================

Private Function DeclineCompoundName(ByVal nameText As String, _
                                     ByVal nameType As NameComponent, _
                                     ByVal gender As GrammaticalGender, _
                                     ByVal targetCase As GrammaticalCase) As String
    Dim parts() As String
    Dim declinedParts() As String
    Dim i As Integer

    parts = Split(nameText, "-")
    ReDim declinedParts(UBound(parts))

    For i = LBound(parts) To UBound(parts)
        If i = UBound(parts) Then
            ' Decline only last part
            declinedParts(i) = DeclineNameInternal(Trim(parts(i)), nameType, gender, targetCase)
        Else
            ' Keep other parts unchanged
            declinedParts(i) = Trim(parts(i))
        End If
    Next i

    DeclineCompoundName = Join(declinedParts, "-")
End Function

'===============================================================================
' HELPER FUNCTIONS
'===============================================================================

Private Function EndsWith(ByVal text As String, ByVal ending As String) As Boolean
    Dim textLen As Long
    Dim endLen As Long

    text = Trim(text)
    ending = Trim(ending)

    textLen = Len(text)
    endLen = Len(ending)

    If textLen < endLen Then
        EndsWith = False
        Exit Function
    End If

    ' Use binary comparison for exact match
    EndsWith = (StrComp(Right(text, endLen), ending, vbBinaryCompare) = 0)
End Function

Private Function IsConsonant(ByVal char As String) As Boolean
    ' Ukrainian consonants using Unicode character codes
    ' б=1073, в=1074, г=1075, ґ=1169, д=1076, ж=1078, з=1079, к=1082, л=1083,
    ' м=1084, н=1085, п=1087, р=1088, с=1089, т=1090, ф=1092, х=1093, ц=1094,
    ' ч=1095, ш=1096, щ=1097, ь=1100

    Dim code As Long

    If Len(char) = 0 Then
        IsConsonant = False
        Exit Function
    End If

    code = AscW(char)

    ' Lowercase Ukrainian consonants
    If code = 1073 Or code = 1074 Or code = 1075 Or code = 1169 Or code = 1076 Or _
       code = 1078 Or code = 1079 Or code = 1082 Or code = 1083 Or code = 1084 Or _
       code = 1085 Or code = 1087 Or code = 1088 Or code = 1089 Or code = 1090 Or _
       code = 1092 Or code = 1093 Or code = 1094 Or code = 1095 Or code = 1096 Or _
       code = 1097 Or code = 1100 Then
        IsConsonant = True
        Exit Function
    End If

    ' Uppercase Ukrainian consonants
    If code = 1041 Or code = 1042 Or code = 1043 Or code = 1168 Or code = 1044 Or _
       code = 1046 Or code = 1047 Or code = 1050 Or code = 1051 Or code = 1052 Or _
       code = 1053 Or code = 1055 Or code = 1056 Or code = 1057 Or code = 1058 Or _
       code = 1060 Or code = 1061 Or code = 1062 Or code = 1063 Or code = 1064 Or _
       code = 1065 Or code = 1068 Then
        IsConsonant = True
        Exit Function
    End If

    IsConsonant = False
End Function