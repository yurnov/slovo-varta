# Slovo-Varta

**Slovo-Varta** is a Google Sheets add-on that provides automatic declension (word inflection) for Ukrainian names. It helps users work with Ukrainian names in different grammatical cases directly within spreadsheets.

## Features

- **Automatic Name Declension**: Decline Ukrainian given names, patronymics, and family names into different grammatical cases
- **Gender Detection**: Automatically determine gender based on name patterns
- **Batch Processing**: Process multiple names at once using spreadsheet formulas
- **Easy Integration**: Works seamlessly within Google Sheets

## Installation

1. Open your Google Sheet
2. Go to **Extensions** → **Add-ons** → **Get add-ons**
3. Search for "Slovo-Varta"
4. Click **Install**
5. Grant necessary permissions

## Usage

### Basic Formula Structure

```
=SlovoVarta(Name, NameType, GrammaticalCase, [Gender])
```

### Parameters

- **Name** (required): The Ukrainian name to decline
- **NameType** (required): Type of name - "GivenName", "Patronymic", or "FamilyName"
- **GrammaticalCase** (required): Target case - "Nominative", "Genitive", "Dative", "Accusative", "Instrumental", "Locative", or "Vocative"
- **Gender** (optional): "Male" or "Female" - if omitted, gender will be auto-detected

### Grammatical Cases

Ukrainian has seven grammatical cases:

1. **Nominative** (називний) - Subject case: Хто? Що?
2. **Genitive** (родовий) - Possessive case: Кого? Чого?
3. **Dative** (давальний) - Indirect object: Кому? Чому?
4. **Accusative** (знахідний) - Direct object: Кого? Що?
5. **Instrumental** (орудний) - Instrument case: Ким? Чим?
6. **Locative** (місцевий) - Location case: На кому? На чому?
7. **Vocative** (кличний) - Address case: direct address

### Examples

#### Example 1: Basic Name Declension

Decline a given name to genitive case:

```
=SlovoVarta("Тарас", "GivenName", "Genitive")
```

Result: `Тараса`

#### Example 2: Patronymic with Gender

Decline a patronymic to dative case with specified gender:

```
=SlovoVarta("Григорович", "Patronymic", "Dative", "Male")
```

Result: `Григоровичу`

#### Example 3: Family Name Auto-Gender

Decline a family name with auto-detected gender:

```
=SlovoVarta("Шевченко", "FamilyName", "Instrumental")
```

Result: `Шевченком` (if male) or `Шевченко` (if female, indeclinable)

## Advanced Usage

### Combining Multiple Functions for Full Names

You can combine multiple `SlovoVarta` functions to decline full names (Given Name + Patronymic + Family Name).

#### Example 1: Dative Case for Certificates

**Template**: Сертифікат виданий [ПІБ в давальному відмінку]

**Input**: Тарас | Григорович | Шевченко

**Formula**:
```
="Сертифікат виданий " & SlovoVarta("Тарас", "GivenName", "Dative") & " " & SlovoVarta("Григорович", "Patronymic", "Dative") & " " & SlovoVarta("Шевченко", "FamilyName", "Dative")
```

**Result**: `Сертифікат виданий Тарасу Григоровичу Шевченку`

#### Example 2: Genitive Case for Orders

**Template**: Призначити на посаду [ПІБ в родовому відмінку]

**Input**: Іван | Якович | Франко

**Formula**:
```
="Призначити на посаду " & SlovoVarta("Іван", "GivenName", "Genitive") & " " & SlovoVarta("Якович", "Patronymic", "Genitive") & " " & SlovoVarta("Франко", "FamilyName", "Genitive")
```

**Result**: `Призначити на посаду Івана Яковича Франка`

#### Example 3: Vocative Case for Letters

**Template**: Шановний/-а [ПІБ в кличному відмінку]!

**Input**: Леся | Петрівна | Українка

**Formula**:
```
="Шановна " & SlovoVarta("Леся", "GivenName", "Vocative") & " " & SlovoVarta("Петрівна", "Patronymic", "Vocative") & " " & SlovoVarta("Українка", "FamilyName", "Vocative") & "!"
```

**Result**: `Шановна Лесю Петрівно Українко!`

### Using with Cell References

Instead of hardcoding names, you can reference cells:

Assuming:
- Cell A1 contains: `Іван`
- Cell B1 contains: `Франко`
- Cell C1 contains: `Genitive`

```
=SlovoVarta(A1, "GivenName", C1)
```

### Batch Processing

You can apply formulas to entire columns for batch processing:

| A | B | C | D |
|---------|--------------|-----------|-------------------|
| Name | Type | Case | Result |
| Олена | GivenName | Dative | =SlovoVarta(A2, B2, C2) |
| Марія | GivenName | Genitive | =SlovoVarta(A3, B3, C3) |
| Андрій | GivenName | Vocative | =SlovoVarta(A4, B4, C4) |

## Name Type Reference

### GivenName (Ім'я)

Ukrainian given names (first names):
- Male: Олександр, Іван, Петро, Андрій, Михайло
- Female: Олена, Марія, Катерина, Наталія, Ірина

### Patronymic (По батькові)

Patronymics formed from father's given name:
- Male endings: -ович, -йович, -ійович (Іванович, Петрович, Сергійович)
- Female endings: -івна, -ївна (Іванівна, Петрівна, Сергіївна)

### FamilyName (Прізвище)

Ukrainian family names:
- Declinable: Шевченко, Коваль, Мельник, Бондар
- Some female forms may be indeclinable depending on ending

## Gender Detection

The add-on automatically detects gender based on:
- Name endings (е.g., -а, -я typically female for given names)
- Patronymic patterns (-ович/-івна)
- Family name patterns

You can override auto-detection by explicitly specifying the Gender parameter.

## Case Usage Guidelines

### Common Use Cases by Grammatical Case

**Nominative (називний)**
- Subject of sentence
- Dictionary form
- Example: **Тарас Шевченко** написав вірш

**Genitive (родовий)**
- Possession
- "Of" constructions
- After numbers, negations
- Example: Твори **Тараса Шевченка**

**Dative (давальний)**
- Indirect object
- Recipient of action
- Example: Дати книгу **Тарасу Шевченку**

**Accusative (знахідний)**
- Direct object
- Example: Я бачу **Тараса Шевченка**

**Instrumental (орудний)**
- Instrument or means
- "With" or "by"
- Professional titles
- Example: Написано **Тарасом Шевченком**

**Locative (місцевий)**
- Location
- "About" or "concerning"
- Example: Думати про **Тараса Шевченка**

**Vocative (кличний)**
- Direct address
- Example: **Тарасе Шевченко**, ти великий поет!

## Supported Names

### Coverage

The add-on supports:
- Common Ukrainian given names (male and female)
- Standard patronymic patterns
- Most Ukrainian family names

### Limitations

- Non-Ukrainian names may not decline correctly
- Rare or archaic names might need manual verification
- Foreign names adapted to Ukrainian may have limited support

## Troubleshooting

### Common Issues

**Error: "Invalid case"**
- Check spelling of the case parameter
- Use English case names: "Nominative", "Genitive", "Dative", etc.

**Unexpected Results**
- Verify the NameType parameter is correct
- Try specifying Gender explicitly
- Check if the name is Ukrainian

**Function Not Found**
- Ensure the add-on is installed and enabled
- Reload the spreadsheet
- Check Extensions → Add-ons → Manage add-ons

### Getting Help

If you encounter issues:
1. Verify all parameters are spelled correctly
2. Check that the name is Ukrainian
3. Try with a common name to test functionality
4. Contact support with specific examples

## Privacy & Data

- All processing happens within Google's infrastructure
- No name data is stored or transmitted to external servers
- The add-on only accesses cells you explicitly reference in formulas

## Technical Details

### Implementation

- Built using Google Apps Script
- Uses rule-based morphological analysis
- Optimized for Ukrainian language patterns

### Performance

- Instant processing for individual names
- Efficient batch processing for large datasets
- No external API calls required

## Examples Library

### Official Documents

**Certificate (Genitive)**
```
Сертифікат виданий [ПІБ в родовому відмінку]
```

**Order (Dative)**
```
Наказ про призначення [ПІБ в давальному відмінку]
```

**Power of Attorney (Genitive)**
```
Довіреність від [ПІБ в родовому відмінку]
```

### Correspondence

**Formal Letter Opening (Dative)**
```
Шановному/-ій [ПІБ в давальному відмінку]
```

**Letter Closing (Genitive)**
```
З повагою, [підпис][ПІБ в родовому відмінку]
```

### Addressing People

**Polite Address (Vocative)**
```
[Ім'я в кличному відмінку], [звернення]
```

## Updates & Changelog

### Version 1.0.0
- Initial release
- Support for 7 grammatical cases
- Three name types (GivenName, Patronymic, FamilyName)
- Automatic gender detection
- Basic error handling

## Contributing

We welcome contributions! If you notice:
- Incorrectly declined names
- Missing name patterns
- Bugs or issues

Please report them through the add-on feedback system.

## License

This add-on is provided as-is for use within Google Sheets.

## About

**Slovo-Varta** (Слово-Варта) - where "Slovo" means "word" and "Varta" means "guard" or "watch" - aims to guard the proper usage of Ukrainian language in documents and correspondence.

---

For more information and updates, visit the add-on page in the Google Workspace Marketplace.