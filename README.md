# Slovo-Varta (Слово-Варта) 🇺🇦

**Ukrainian Name Declension for Microsoft Excel**

[![License:  MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![VBA](https://img.shields.io/badge/VBA-Excel-green.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![Ukrainian](https://img.shields.io/badge/Language-Ukrainian-yellow.svg)](https://en.wikipedia.org/wiki/Ukrainian_language)

Slovo-Varta is an open-source VBA module for Microsoft Excel designed for the automatic declension of Ukrainian names, surnames, and patronymics (відмінювання українських імен, прізвищ та по батькові).

---

## 📋 Table of Contents

- [Purpose](#-purpose)
- [Features](#-features)
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Usage Examples](#-usage-examples)
- [Function Reference](#-function-reference)
- [Supported Cases](#-supported-cases)
- [Limitations](#-limitations)
- [Contributing](#-contributing)
- [License](#-license)
- [Acknowledgments](#-acknowledgments)
- [Support the Project](#-support-the-project)

---

## 🎯 Purpose

Administrative work in the Ukrainian military and public sector often requires processing **thousands of names** in different grammatical cases. Whether it's generating certificates, orders, diplomas, or official documents, manually declining names is: 

- ⏰ **Time-consuming** - Hours wasted on repetitive work
- ❌ **Error-prone** - Manual mistakes in official documents
- 📊 **Inefficient** - Takes focus away from critical tasks

**Slovo-Varta automates this process**, reducing manual errors and saving time for more important work. 

### Real-World Use Cases

- 📜 **Military Orders** - "Призначити на посаду [ПІБ в родовому відмінку]"
- 🎓 **Diplomas & Certificates** - "Видано [ПІБ в давальному відмінку]"
- 📝 **Official Documents** - Automated name processing for thousands of personnel
- 📧 **Correspondence** - "Шановному/Шановній [ПІБ в давальному]"

---

## ✨ Features

- ✅ **Genitive Case** (Родовий відмінок) - кого?  чого?
- ✅ **Dative Case** (Давальний відмінок) - кому? чому?
- ✅ **Given Names** (Ім'я) - Тарас → Тараса, Тарасу
- ✅ **Patronymics** (По батькові) - Григорович → Григоровича, Григоровичу
- ✅ **Family Names** (Прізвище) - Шевченко → Шевченка, Шевченку
- ✅ **Compound Names** (Складні імена) - Нечуй-Левицький → Нечуя-Левицького
- ✅ **Adjective Surnames** - Новоставський → Новоставського, Новоставському
- ✅ **Gender Support** - Multiple formats:  m/f, ч/ж, masculine/feminine
- ✅ **Excel Functions** - Easy-to-use formulas like `=GivenNameGenitive("Тарас", "m")`
- ✅ **No External Dependencies** - Pure VBA, works offline

---

## 🔧 Installation

### Step 1: Enable Developer Tab (if not visible)

1. Open Excel
2. Go to **File** → **Options** → **Customize Ribbon**
3. Check ✅ **Developer**
4. Click **OK**

### Step 2: Import the VBA Module

1. Open your Excel file
2. Press **Alt + F11** (Windows) or **Fn + Option + F11** (Mac) to open VBA Editor
3. In the menu, click **File** → **Import File.. .**
4. Select the `SlovoVarta.bas` file
5. Press **Ctrl + S** to save
6. Close VBA Editor
7. **Save your file as `.xlsm`** (Excel Macro-Enabled Workbook)

### Step 3: Enable Macros

1. When opening the file, click **Enable Content** in the yellow security bar
2. Or:  **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings** → Select "Enable all macros"

**Alternative:  Manual Import**

If you prefer to copy-paste: 

1. Open VBA Editor (**Alt + F11**)
2. Click **Insert** → **Module**
3. Copy the entire contents of `SlovoVarta.bas`
4. Paste into the module window
5. Save as `.xlsm`

---

## 🚀 Quick Start

### Example 1: Genitive Case for Certificates

Create a certificate:  "Сертифікат виданий [ПІБ в родовому відмінку]"

| A | B | C | D | E |
|---|---|---|---|---|
| **Ім'я** | **По батькові** | **Прізвище** | **Стать** | **Сертифікат** |
| Тарас | Григорович | Шевченко | m | =CONCATENATE("Сертифікат виданий ", GivenNameGenitive(A2,$D2), " ", PatronymicGenitive(B2,$D2), " ", FamilyNameGenitive(C2,$D2)) |

**Result:**  
`Сертифікат виданий Тараса Григоровича Шевченка`

### Example 2: Dative Case for Orders

Military order: "Призначити на посаду [ПІБ в давальному відмінку]"

| A | B | C | D | E |
|---|---|---|---|---|
| **Ім'я** | **По батькові** | **Прізвище** | **Стать** | **Наказ** |
| Юрій | Ігорович | Новоставський | m | =CONCATENATE("Призначити на посаду ", GivenNameDative(A2,$D2), " ", PatronymicDative(B2,$D2), " ", FamilyNameDative(C2,$D2)) |

**Result:**  
`Призначити на посаду Юрію Ігоровичу Новоставському`

---

## 📖 Usage Examples

### Basic Functions

```excel
' Given Name (Ім'я)
=GivenNameGenitive("Тарас", "m")      → "Тараса"
=GivenNameDative("Тарас", "m")        → "Тарасу"

' Patronymic (По батькові)
=PatronymicGenitive("Григорович", "m") → "Григоровича"
=PatronymicDative("Григорович", "m")   → "Григоровичу"

' Family Name (Прізвище)
=FamilyNameGenitive("Шевченко", "m")  → "Шевченка"
=FamilyNameDative("Шевченко", "m")    → "Шевченку"
```

### Universal Function

```excel
=DeclineName("Шевченко", "family", "m", "genitive") → "Шевченка"
=DeclineName("Людмила", "given", "f", "dative")     → "Людмилі"
```

### Gender Formats

All these formats work: 

```excel
=GivenNameGenitive("Тарас", "m")          ✅
=GivenNameGenitive("Тарас", "ч")          ✅
=GivenNameGenitive("Тарас", "masculine")  ✅
=GivenNameGenitive("Тарас", "чоловік")    ✅

=GivenNameGenitive("Марія", "f")          ✅
=GivenNameGenitive("Марія", "ж")          ✅
=GivenNameGenitive("Марія", "feminine")   ✅
=GivenNameGenitive("Марія", "жінка")      ✅
```

### Batch Processing

Process entire columns:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| **Ім'я** | **По батькові** | **Прізвище** | **Стать** | **Ім'я (Р. в.)** | **По батькові (Р.в.)** | **Прізвище (Р.в.)** |
| Тарас | Григорович | Шевченко | m | `=GivenNameGenitive(A2,$D2)` | `=PatronymicGenitive(B2,$D2)` | `=FamilyNameGenitive(C2,$D2)` |
| Леся | Петрівна | Українка | f | `=GivenNameGenitive(A3,$D3)` | `=PatronymicGenitive(B3,$D3)` | `=FamilyNameGenitive(C3,$D3)` |
| Іван | Якович | Франко | m | `=GivenNameGenitive(A4,$D4)` | `=PatronymicGenitive(B4,$D4)` | `=FamilyNameGenitive(C4,$D4)` |

**Tip:** Use `$D2` (absolute column reference) for gender so it doesn't change when copying formulas.

---

## 📚 Function Reference

### Main Functions

#### `GivenNameGenitive(givenName, gender)`
Decline given name (ім'я) to genitive case (родовий відмінок).

**Parameters:**
- `givenName` (String) - Given name in nominative case
- `gender` (String) - Gender: "m"/"f"/"ч"/"ж"/"masculine"/"feminine"

**Returns:** String - Declined given name

**Example:**
```excel
=GivenNameGenitive("Юрій", "m") → "Юрія"
```

---

#### `GivenNameDative(givenName, gender)`
Decline given name (ім'я) to dative case (давальний відмінок).

**Example:**
```excel
=GivenNameDative("Юрій", "m") → "Юрію"
```

---

#### `PatronymicGenitive(patronymic, gender)`
Decline patronymic (по батькові) to genitive case.

**Example:**
```excel
=PatronymicGenitive("Ігорович", "m") → "Ігоровича"
```

---

#### `PatronymicDative(patronymic, gender)`
Decline patronymic (по батькові) to dative case.

**Example:**
```excel
=PatronymicDative("Ігорович", "m") → "Ігоровичу"
```

---

#### `FamilyNameGenitive(familyName, gender)`
Decline family name (прізвище) to genitive case.

**Example:**
```excel
=FamilyNameGenitive("Новоставський", "m") → "Новоставського"
```

---

#### `FamilyNameDative(familyName, gender)`
Decline family name (прізвище) to dative case.

**Example:**
```excel
=FamilyNameDative("Новоставський", "m") → "Новоставському"
```

---

### Universal Function

#### `DeclineName(nameText, nameType, gender, targetCase)`
Universal function for declining any name component.

**Parameters:**
- `nameText` (String) - Name in nominative case
- `nameType` (String) - Type: "given"/"patronymic"/"family"
- `gender` (String) - Gender: "m"/"f"/"ч"/"ж"
- `targetCase` (String) - Case: "genitive"/"dative"

**Example:**
```excel
=DeclineName("Шевченко", "family", "m", "genitive") → "Шевченка"
```

---

### Utility Functions

#### `DebugDecline(nameText, nameType, gender, targetCase)`
Debug function showing detailed declension process.

**Example:**
```excel
=DebugDecline("Юрій", "given", "m", "dative")
```

Returns detailed debug information for troubleshooting.

---

#### `SlovoVartaVersion()`
Returns version information. 

**Example:**
```excel
=SlovoVartaVersion()
→ "Slovo-Varta v1.0.0 - Ukrainian Name Declension for Excel"
```

---

## 📖 Supported Cases

### Genitive Case (Родовий відмінок)
**Question:** Кого? Чого?  (Of whom? Of what?)

**Usage:**
- Possession: "книга **Тараса**" (Taras's book)
- After numbers: "п'ять **студентів**"
- After "немає": "немає **Марії**"
- Certificates: "Сертифікат виданий **Тараса Григоровича Шевченка**"

**Examples:**
| Nominative | Genitive |
|------------|----------|
| Тарас | Тараса |
| Марія | Марії |
| Шевченко | Шевченка |

---

### Dative Case (Давальний відмінок)
**Question:** Кому? Чому? (To whom? To what?)

**Usage:**
- Indirect object: "дати **Іванові**" (give to Ivan)
- Orders: "Призначити на посаду **Петру Івановичу Сидоренку**"
- Certificates: "Видано **Марії Петрівні Коваленко**"
- Age: "**Марії** 25 років"

**Examples:**
| Nominative | Dative |
|------------|--------|
| Тарас | Тарасу |
| Марія | Марії |
| Шевченко | Шевченку |

---

## ⚠️ Limitations

### Currently Not Supported

- ❌ **Accusative case** (Знахідний) - кого? що? 
- ❌ **Ablative case** (Орудний) - ким?  чим?
- ❌ **Locative case** (Місцевий) - на кому? на чому?
- ❌ **Vocative case** (Кличний) - direct address
- ❌ **Automatic gender detection** - gender must be specified
- ❌ **Plural forms** - only singular names

### Edge Cases

- Some **foreign names** may not decline correctly
- **Historical or rare names** might need manual adjustment
- Compound names with **more than 2 parts** might have issues

### Known Issues

If you encounter issues, please: 
1. Check the examples in this README
2. Use the `DebugDecline()` function to diagnose
3. [Open an issue](https://github.com/yurnov/slovo-varta/issues) on GitHub

---

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Report Bugs** - [Open an issue](https://github.com/yurnov/slovo-varta/issues)
3. **Submit Pull Requests** - Add support for new name patterns
4. **Improve Documentation** - Help make the README clearer
5. **Test Edge Cases** - Report names that don't decline correctly

### Development

Created with support of **GitHub Copilot** using **Claude Sonnet 4.5** model.

---

## 📄 License

This project is licensed under the **[MIT License](LICENSE)**.

---

## 🙏 Acknowledgments

### Inspired By

This project was inspired by the excellent **[shevchenko-js](https://github.com/tooleks/shevchenko-js)** library by [tooleks](https://github.com/tooleks). Shevchenko-js provides comprehensive Ukrainian name declension for JavaScript/TypeScript applications.  If you need a solution for web or Node.js, check it out!

### Special Thanks

- **Authors of [shevchenko-js](https://github.com/tooleks/shevchenko-js)** - for the inspiration and linguistic foundation
- **Defense Forces of Ukraine** (Сили оборони України) 🇺🇦 - for defending our homeland
<!-- - **All contributors** - for making this project better -->

---

## 💙💛 Support the Project

If you find **Slovo-Varta** helpful, the best way to say "thank you" is to **donate** to: 

### **Come Back Alive Foundation** (Повернись живим)
**[🔗 Donate Here](https://savelife.in. ua/en/donate-en)**

Come Back Alive is a charitable foundation that comprehensively equips the Defence Forces of Ukraine with: 
- 🚁 Drones and UAV systems
- 🎯 Tactical gear and communication systems
- 📡 Electronic warfare equipment
- 🎓 Educational programs for the military
- and much more

**Every donation helps protect Ukraine and save lives. ** 🇺🇦

---

## 🌟 Star the Project

If you find this project useful, please give it a ⭐ on GitHub!

---

**Slava Ukraini! ** 🇺🇦 **Героям слава!**

---

## 📈 Changelog

### Initial version
- ✅ Initial release
- ✅ Genitive and dative case support
- ✅ Given names, patronymics, and family names
- ✅ Multiple gender format support
- ✅ Compound name handling
- ✅ Adjective surname support

---

**Made with 💙💛 for Ukraine**