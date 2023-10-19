# CustomVStack AppScript Function

This script provides a custom functionality to vertically stack values from a given range in Google Sheets.

Appends ranges vertically and in sequence to return a larger array.

## Usage

### Function

```javascript
CustomVStack(range, skipEmpty)
```

### Parameters

- range (${\color{lime}required}$): Range of the data set (e.g., "A1:C10")

  ${\color{lightblue}Input}$ : `string`

    

- skipEmpty (${\color{orange}optional}$): Whether to skip the empty values or not
    
    ${\color{lightblue}Input}$ : `boolean`
    
    ${\color{lightgreen}Default}$ : `True`
    
> [!NOTE]
> The range parameter must be provided as a string within the formula and USES the active sheet ONLY.

### Example

```
=CustomVStack("A1:C10", True)
```
Or,
```
=CustomVStack("A1:C10")
```