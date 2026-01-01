# 🇺🇦 SlovoVarta (СловоВарта)

**Ukrainian Name Declension Module for Microsoft Excel**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Language: VBA](https://img.shields.io/badge/Language-VBA-blue.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![Excel Version](https://img.shields.io/badge/Excel-2010%2B-green.svg)](https://www.microsoft.com/en-us/microsoft-365/excel)

> **SlovoVarta** (СловоВарта) - від "слово" (word) та "варта" (guard/protector). A tool that protects the proper declension of Ukrainian names.

## 📋 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Installation](#-installation)
- [Usage](#-usage)
  - [Excel Functions](#excel-functions)
  - [VBA Procedures](#vba-procedures)
- [Examples](#-examples)
- [Technical Details](#-technical-details)
- [Known Issues](#-known-issues)
- [Contributing](#-contributing)
- [License](#-license)
- [Support](#-support)

## 🎯 Overview

SlovoVarta is a VBA module that enables proper grammatical declension of Ukrainian first names, patronymics, and surnames directly in Microsoft Excel. It handles the complexities of Ukrainian grammar, including:

- Six grammatical cases (nominative, genitive, dative, accusative, instrumental, locative)
- Gender-specific declension rules
- Special cases and exceptions
- Compound names and names with hyphens
- Foreign names that follow Ukrainian declension patterns

## ✨ Features

- **📊 Excel Functions** - Use Ukrainian name declension directly in Excel formulas
- **🔧 VBA API** - Integrate declension into your VBA macros and applications
- **👥 Gender Support** - Handles masculine, feminine, and neutral gender names
- **📝 All Six Cases** - Complete support for all Ukrainian grammatical cases
- **🎭 Name Types** - Works with first names (given names), patronymics, and surnames
- **🌐 Unicode Support** - Full support for Ukrainian characters
- **⚡ Performance** - Optimized for large datasets
- **🛡️ Error Handling** - Graceful handling of edge cases and invalid inputs

## 📥 Installation

### Method 1: Import the .BAS Module (Recommended)

1. Download the `SlovoVarta.bas` file from this repository
2. Open your Excel workbook
3. Press `Alt + F11` to open the VBA Editor
4. Go to **File** → **Import File** (or press `Ctrl + M`)
5. Select the downloaded `SlovoVarta.bas` file
6. Click **Open**
7. Save your workbook as `.xlsm` (Excel Macro-Enabled Workbook)

### Method 2: GitHub Import

1. In Excel, press `Alt + F11` to open VBA Editor
2. Go to **File** → **Import File**
3. Navigate to the repository location and select `SlovoVarta.bas`
4. Save workbook as `.xlsm`

### Method 3: Direct Copy-Paste

1. Open the `SlovoVarta.bas` file in a text editor
2. Copy all content
3. In Excel, press `Alt + F11`
4. Click **Insert** → **Module**
5. Paste the code
6. Save workbook as `.xlsm`

**Alternative: Manual Import**

If you prefer to manually import the code:
1. Create a new module in your VBA project
2. Copy the contents of `SlovoVarta.bas`
3. Paste into the new module

### ⚠️ Known Encoding Issue

When importing the `SlovoVarta.bas` file into Excel VBA Editor, you may encounter **incorrect character encoding** for Ukrainian text in comments and string literals. The text may appear as garbled characters (e.g., `Ð'Ñ–Ð²Ñ‡Ð°Ð»Ð¾` instead of `Вівчарь`).

![image1](image1)

**Important:** Even with this visual encoding issue, **the module still works correctly** because:
- The actual string processing uses `ChrW()` function with Unicode code points
- All Ukrainian characters are represented as Unicode values, not as literal characters
- The encoding issue only affects human-readable comments and examples in the code

#### Recommended Solutions:

**Option 1: Set Windows Regional Settings (Preferred)**
1. Open **Control Panel** → **Region** (or **Clock and Region** → **Region**)
2. Click **Administrative** tab
3. Click **Change system locale...**
4. Select **Ukrainian (Ukraine)** or ensure **Beta: Use Unicode UTF-8 for worldwide language support** is checked
5. Click **OK** and restart your computer
6. Re-import the `SlovoVarta.bas` file

**Option 2: Use Manual Import Method**
1. Open the `SlovoVarta.bas` file in a UTF-8 compatible editor (e.g., Visual Studio Code, Notepad++)
2. Ensure the file is opened with **UTF-8** encoding
3. Copy all contents
4. In Excel VBA Editor (**Alt + F11**), click **Insert** → **Module**
5. Paste the code into the new module
6. Save as `.xlsm`

**Option 3: Live with the Visual Issue**
- If changing system settings is not an option, you can use the module as-is
- The garbled text in comments does not affect functionality
- All Excel functions will work correctly with Ukrainian names

#### Verification:
To verify the module works correctly regardless of the encoding display issue, test with:
```excel
=GivenNameGenitive("Тарас", "m")
```
Expected result: `Тараса`

If the function returns the correct result, the module is working properly.

## 🚀 Quick Start

```excel
' Genitive case (родовий відмінок)
=GivenNameGenitive("Тарас", "m")      ' Returns: Тараса
=SurnameGenitive("Шевченко", "m")     ' Returns: Шевченка

' Dative case (давальний відмінок)
=GivenNameDative("Олена", "f")        ' Returns: Олені
=PatronymicDative("Петрівна", "f")    ' Returns: Петрівні

' Full name declension
=FullNameGenitive("Тарас", "Григорович", "Шевченко", "m")
' Returns: Тараса Григоровича Шевченка
```

## 📖 Usage

### Excel Functions

The module provides Excel functions for each grammatical case and name type:

#### Given Names (First Names)

| Function | Case | Example Input | Example Output |
|----------|------|---------------|----------------|
| `GivenNameGenitive(name, gender)` | Genitive | Іван, m | Івана |
| `GivenNameDative(name, gender)` | Dative | Марія, f | Марії |
| `GivenNameAccusative(name, gender)` | Accusative | Олександр, m | Олександра |
| `GivenNameInstrumental(name, gender)` | Instrumental | Катерина, f | Катериною |
| `GivenNameLocative(name, gender)` | Locative | Петро, m | Петрові |
| `GivenNameVocative(name, gender)` | Vocative | Андрій, m | Андрію |

#### Patronymics

| Function | Case | Example Input | Example Output |
|----------|------|---------------|----------------|
| `PatronymicGenitive(patronymic, gender)` | Genitive | Іванович, m | Івановича |
| `PatronymicDative(patronymic, gender)` | Dative | Петрівна, f | Петрівні |
| `PatronymicAccusative(patronymic, gender)` | Accusative | Миколайович, m | Миколайовича |
| `PatronymicInstrumental(patronymic, gender)` | Instrumental | Олександрівна, f | Олександрівною |
| `PatronymicLocative(patronymic, gender)` | Locative | Васильович, m | Васильовичу |
| `PatronymicVocative(patronymic, gender)` | Vocative | Григорівна, f | Григорівно |

#### Surnames

| Function | Case | Example Input | Example Output |
|----------|------|---------------|----------------|
| `SurnameGenitive(surname, gender)` | Genitive | Шевченко, m | Шевченка |
| `SurnameDative(surname, gender)` | Dative | Коваленко, f | Коваленко |
| `SurnameAccusative(surname, gender)` | Accusative | Мельник, m | Мельника |
| `SurnameInstrumental(surname, gender)` | Instrumental | Бондар, m | Бондарем |
| `SurnameLocative(surname, gender)` | Locative | Ткач, m | Ткачу |
| `SurnameVocative(surname, gender)` | Vocative | Коваль, m | Ковалю |

#### Full Names

| Function | Case | Parameters |
|----------|------|------------|
| `FullNameGenitive(given, patronymic, surname, gender)` | Genitive | All name parts |
| `FullNameDative(given, patronymic, surname, gender)` | Dative | All name parts |
| `FullNameAccusative(given, patronymic, surname, gender)` | Accusative | All name parts |
| `FullNameInstrumental(given, patronymic, surname, gender)` | Instrumental | All name parts |
| `FullNameLocative(given, patronymic, surname, gender)` | Locative | All name parts |
| `FullNameVocative(given, patronymic, surname, gender)` | Vocative | All name parts |

**Parameters:**
- `name` / `given` / `patronymic` / `surname` - Ukrainian name (String)
- `gender` - Gender: "m" (masculine), "f" (feminine), or "n" (neutral) (String)

### VBA Procedures

For VBA integration, use the core functions:

```vba
' Core declension function
Function DeclineUkrainianName(name As String, gender As String, nameType As String, grammaticalCase As String) As String

' Parameters:
' - name: The Ukrainian name to decline
' - gender: "m" (masculine), "f" (feminine), "n" (neutral)
' - nameType: "given" (first name), "patronymic", "surname"
' - grammaticalCase: "genitive", "dative", "accusative", "instrumental", "locative", "vocative"

' Example:
Dim declined As String
declined = DeclineUkrainianName("Іван", "m", "given", "genitive")
' Returns: Івана
```

## 💡 Examples

### Basic Usage

```excel
' Single names
=GivenNameGenitive("Богдан", "m")           ' → Богдана
=PatronymicDative("Михайлівна", "f")        ' → Михайлівні
=SurnameInstrumental("Коваленко", "m")      ' → Коваленком

' Full names
=FullNameGenitive("Леся", "Петрівна", "Українка", "f")
' → Лесі Петрівни Українки

=FullNameDative("Тарас", "Григорович", "Шевченко", "m")
' → Тарасу Григоровичу Шевченку
```

### Advanced Examples

```excel
' Hyphenated names
=GivenNameGenitive("Анна-Марія", "f")       ' → Анни-Марії

' Names ending with special characters
=SurnameGenitive("Савченко", "m")           ' → Савченка
=SurnameGenitive("Савченко", "f")           ' → Савченко (no declension for feminine -енко surnames)

' Foreign names adapted to Ukrainian
=GivenNameGenitive("Джон", "m")             ' → Джона
=SurnameGenitive("Сміт", "m")               ' → Сміта
```

### Batch Processing

```vba
Sub DeclineNamesList()
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        ' Assuming columns: A=FirstName, B=Patronymic, C=Surname, D=Gender
        ' Output in column E
        Cells(i, "E").Value = FullNameGenitive( _
            Cells(i, "A").Value, _
            Cells(i, "B").Value, _
            Cells(i, "C").Value, _
            Cells(i, "D").Value _
        )
    Next i
End Sub
```

## 🔧 Technical Details

### Grammatical Cases

Ukrainian has six grammatical cases, each with specific usage:

| Case | Ukrainian | Question | Usage Example |
|------|-----------|----------|---------------|
| Nominative | Називний | Хто? Що? | Іван пише листа |
| Genitive | Родовий | Кого? Чого? | Книга Івана |
| Dative | Давальний | Кому? Чому? | Дати Іванові |
| Accusative | Знахідний | Кого? Що? | Бачу Івана |
| Instrumental | Орудний | Ким? Чим? | З Іваном |
| Locative | Місцевий | На кому? На чому? | Про Івана |
| Vocative | Кличний | - | Іване! |

### Declension Rules

The module implements Ukrainian grammatical rules for:

1. **Given Names** - Based on ending patterns and gender
   - Masculine: -о → -а, consonant → +а, -ій → -ія, etc.
   - Feminine: -а → -и/-і, -я → -і, consonant → no change, etc.

2. **Patronymics** - Regular patterns for -ович/-івна suffixes
   - Masculine: -ович, -євич, -їч
   - Feminine: -івна, -ївна

3. **Surnames** - Complex rules based on endings and gender
   - Declinable: -енко, -ук, -юк, -ський, consonants, etc.
   - Non-declinable: Some -енко for feminine, foreign names, etc.

### Gender Specification

- `"m"` - Masculine (чоловічий рід)
- `"f"` - Feminine (жіночий рід)
- `"n"` - Neutral (середній рід) - rare for personal names

### Performance Considerations

- String operations are optimized for Ukrainian Unicode characters
- Function caching can be implemented for repeated calls
- Handles datasets with thousands of names efficiently

## ⚠️ Known Issues

- **Limited exception handling** - Some rare or non-standard names may not decline correctly
- **Foreign names** - Names that don't follow Ukrainian phonetic patterns may have unexpected results
- **Compound surnames** - Double-barreled surnames may require manual handling
- **Historical names** - Old Ukrainian names may use different declension patterns
- **Character encoding in VBA Editor** - When importing the .BAS file, Ukrainian text in comments may appear garbled due to system locale settings. This is a visual issue only and does not affect functionality. See [Installation](#-installation) section for solutions.

**Recommendation:** Always verify declensions for critical applications, especially for uncommon names.

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Report Issues** - Found a name that doesn't decline correctly? Open an issue!
2. **Suggest Improvements** - Have ideas for better algorithms? Share them!
3. **Add Test Cases** - Help expand the test coverage
4. **Documentation** - Improve examples and explanations

### Development Guidelines

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Make your changes
4. Test thoroughly with various Ukrainian names
5. Commit your changes (`git commit -am 'Add new feature'`)
6. Push to the branch (`git push origin feature/improvement`)
7. Create a Pull Request

### Testing

When contributing, please test your changes with:
- Common Ukrainian names
- Edge cases (hyphenated names, foreign names)
- All grammatical cases
- Both masculine and feminine genders

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2024 SlovoVarta Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 💬 Support

### Questions?

- 📖 Check the [Examples](#-examples) section
- 🐛 [Open an issue](../../issues) for bugs
- 💡 [Start a discussion](../../discussions) for questions

### Resources

- [Ukrainian Grammar Reference](https://uk.wikipedia.org/wiki/Відмінювання_в_українській_мові)
- [Ukrainian Language Rules](http://www.pravopys.net/)
- [Excel VBA Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)

---

**Made with 💙💛 for Ukrainian language preservation**

*SlovoVarta - Protecting Ukrainian words, one declension at a time.*
