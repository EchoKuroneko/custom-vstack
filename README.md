# CustomVStack AppScript Function

This script provides a custom functionality to vertically stack values from a given range in Google Sheets.

Appends ranges vertically and in sequence to return a larger array.

## Purpose

This custom function was created to simplify the process of vertically stacking values from multiple columns in Google Sheets. When dealing with large datasets spanning multiple columns, manually writing a lengthy formula for each column can be time-consuming and error-prone. By using the `CustomVStack` function, large datasets can be combined from various columns into a single vertically stacked array, enhancing efficiency and streamlining data manipulation tasks.

### Built-in Vstack Formula

  ```
  =VSTACK(A1:A10, B1:B10, C1:C10)
  ```

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