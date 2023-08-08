# Class Fmt
Fmt class is a predecalred Id with the default memeber( deb) being the default member returning a new instance of the predeclaredId.
The Fmt class allows a limited form of string interpolation where the locations for substitution in a string are defined by '{}'

The substition field may indicate formatting characters or a variable.  

- Variable fields  {number} where number is the left to right position in the variables list (Zero based indexing)
- Formatting characters {formatting string} : newline, tab, quotation characters.

Formating character instructions may optionally be followed by a number which indicates the number of times to insert that character.  Only one formatting character per field is allowed.

Text output is provided by two methods
- Fmt.Dbg Template, list of variables : Sends the formatted text to Print.Debug and also returns the string
- Fmt.Text(template, list of variables) : returns the formatted string

```
  Fmt.Dbg "{0}{nl}{1}{nl3}{2}(3)", "Hello", "There","World"
' produces the debug output of
Hello
There

World{3}
' there is no variable 3 so the field is just printed
```
Both Fmt.Dbg and Fmt.Text implement the ArrayOp.Splat method, meaning that if a single variable is provided which is a one dimensional array, then this array will be interpreted as if it were a paramarray.
```
Fmt.Dbg "{0}{nl}{1}{nl3}{2}", Array("Hello","There","World")
' gives the output
Hello
There

World
' to print a single one dimensions array as an array encapsulate it in an array
Fmt.Dbg "{0}{nl}{1}{nl3}{2}", Array(Array("Hello","There","World"))
' gives the output
[Hello,There,World]
{1}

{2}
```
## Formatting characters
The available formating characters are
- nl: new line (vbCrLf)
- tb: Tab
- nt: 1 or more newlines followed by one Tab
- dq: Double quote "
- sq: Single quote '
- so: Smart quote Open ‘ ALT+0145
- sc: Smart quote close ’ ALT+0146
- do: Smart Double quote open “ ALT+0147
- dc: Smart Double quote close ” ALT+0148

## What gets printed
Fmt.Dbg and Fmt.Text have the capability of printing container objects and some form of representation of objects
- Arrays A single linear output enclosed by [] with items separated by ",".  An update for multidimensional arrays is planned
- Lists: A single linear output enclosed by {} with items separated by ",".
- Dictionaries A single linear output enclosed by {} with items separated by ",".  The keys are enclosed by double quotes and the key and Item separated by ": "
- Admin:  The textual representation of the admin item is provided
- ItemObjects: The value provided is determnned by a list of options
  - Output of default method if parameterless
  -  An attempt at calling the following functions or properties on the object
     - ToString
     - ToJson
     - Value
     - Name
     - Typename  
     An update is planned to allow a user defined method name to be provided.
Markup is applied recursively, so if a dictionary contains seq items, then the seq items are printed as seq objects
### Formatting Markup
Formatting Markup may be defined by the user using the following methods
- ResetMarkup
  SetArrayMarkup 
- SetDictionaryItemMarkup
- SetItemSeparator
- SetNoMarkup
- SetObjectMarkup

The set methods provide default values if called with no parameters.
```
  Fmt.SetArrayMarkup.SetObjectMarkup("<<<","?",">>>".SetItemSeparator.Dbg "{Fmt.Dbg "{0}{nl}{1}{nl3}{2}", "Hello", Array("There","World"),SeqA("Its","A","Nice","Day")
```       
### Formatting within variable fields
The appearance of text withing variable fields, number of decimal places, currency etc, is devolved to the VBA Format method applied to the variables in the variables list. 

