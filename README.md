<div align="center">

## Bit IO


</div>

### Description

This module allows you to view a file as a collection of bits rather than as a collection of bytes. It allows you to read/write a single bit at a time or read/write up to 32 bits at once.
 
### More Info
 
It's all explained in the code

Same as above

Don't try writing to a file opened for reading, and don't try reading from a file opened for writing - there is no error checking for that and the results are unpredictable.

Don't try to read or write more than 32 bits at a time with the InputBits and OutputBits functions.

If you try to write a value with less bits than that value requires, the correct value will not be written. For example, don't try to write the value 32 into a file using only 4 bits.

After every call to inputbits and inputbit, you should check for eof on the input file using this code:

'inputbit/inputbits call here

if eof(bitfile.filenum) = True then 'replace bitfile with the name of the variable

'put code to exit loop or leave function here

end if


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Derek Haas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/derek-haas.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/derek-haas-bit-io__1-2903/archive/master.zip)





### Source Code

```
Type BitFile
  FileNum As Integer 'File handle
  holder As Byte   'holds a byte from file
  mask As Byte    'used to read bits
End Type
Public Function OpenOBitFile(FileName As String) As BitFile
'Parameters - Filename
'Returns - Bitfile
'What it does - Opens a file for output a single bit at a time
'Example -  dim OutputFile as bitfile
'      OutputFile = OpenOBitFile("C:\test.bit")
 Dim bitfilename As BitFile
  FileNum = FreeFile             'get lowest available file handle
  Open FileName For Binary As FileNum     'open it
  bitfilename.FileNum = FileNum        'assign file number to structure
  bitfilename.holder = 0           'bit holder = 0
  bitfilename.mask = 128           'used to read individual bits
  OpenOBitFile = bitfilename
End Function
Public Function OpenIBitFile(FileName As String) As BitFile
'Parameters - Filename
'Returns - Bitfile
'What it does - Opens a file for input a single bit at a time
'Example -  dim InputFile as bitfile
'      InputFile = OpenIBitFile("C:\command.com")
  Dim bitfilename As BitFile
  FileNum = FreeFile             'get lowest available file handle
  Open FileName For Binary As FileNum     'open it
  bitfilename.FileNum = FileNum        'assign file number to structure
  bitfilename.holder = 0           'bit holder = 0
  bitfilename.mask = 128           'used to read individual bits
  OpenIBitFile = bitfilename
End Function
Public Sub CloseIBitFile(bitfilename As BitFile)
'Parameters - bitfile
'Returns - Nothing
'What it does - Closes the file associated with a bitfile
'Example - CloseIBitFile(InputFile)
  Close bitfilename.FileNum          'Close the file associated with the bitfile
End Sub
Public Sub CloseOBitFile(bitfilename As BitFile)
'Parameters - bitfile
'Returns - Nothing
'What it does - Closes the file associated with a bitfile
'Example - CloseOBitFile(OutputFile)
  If bitfilename.mask <> 128 Then    'If there is unwritten data...
    Put bitfilename.FileNum, , bitfilename.holder  'Write it now
  End If
  Close bitfilename.FileNum    'Close the file
End Sub
Public Sub OutputBit(ByRef bitfilename As BitFile, bit As Byte)
'Parameters - bitfile, bit to write
'Returns - nothing
'What it does - Writes the specified bit to the file
'Example - OutputBit(OutputFile, 1)
  If bit <> 0 Then
    bitfilename.holder = bitfilename.holder Or bitfilename.mask
    'the holder stores up written bits until there are 8
    'At that point vb's normal file handling facilities can write it
  End If
  bitfilename.mask = bitfilename.mask \ 2 'decrease mask by power of 2
  If bitfilename.mask = 0 Then           'if mask is empty
    Put bitfilename.FileNum, , bitfilename.holder 'write the byte
    bitfilename.holder = 0            'reset holder and mask
    bitfilename.mask = 128
  End If
End Sub
Public Sub OutputBits(ByRef bitfilename As BitFile, ByVal code As Long, ByVal count As Integer)
'Parameters - bitfile, data to write, number of bits to use
'Returns - nothing
'What it does - Writes the specified info using the specified number of bits
'Example - OutputBits(OutputFile, 28, 7)
  Dim mask As Long
  mask = 2 ^ (count - 1)
  Do While mask <> 0
    If (mask And code) <> 0 Then      'if the bits match up...
      bitfilename.holder = bitfilename.holder Or bitfilename.mask 'put the bit in the holder
    End If
    bitfilename.mask = bitfilename.mask \ 2
    mask = mask \ 2
    If bitfilename.mask = 0 Then    'when there are 8 bits, write the holder to the file
      Put bitfilename.FileNum, , bitfilename.holder
      bitfilename.holder = 0     'and reset the holder and mask
      bitfilename.mask = 128
    End If
  Loop
End Sub
Public Function InputBit(ByRef bitfilename As BitFile) As Byte
'Parameters - bitfile
'returns - the next bit from the file
'Example: bit = InputBit(InputBitFile)
  Dim value As Byte
  If bitfilename.mask = 128 Then           'if at end of previous byte
    Get bitfilename.FileNum, , bitfilename.holder  'get a new byte from file
  End If
  value = bitfilename.holder And bitfilename.mask   'get the bit
  bitfilename.mask = bitfilename.mask \ 2       'move the mask bit down one
  If bitfilename.mask = 0 Then
    bitfilename.mask = 128
  End If
  If value <> 0 Then                 'return 0 or 1 depending on value
    InputBit = 1
  Else
    InputBit = 0
  End If
End Function
Public Function InputBits(ByRef bitfilename As BitFile, count As Integer) As Long
'Parameters - bitfile, number of bits to read
'returns - the value of the next count bits in the bitfile
'Example: byte = InputBits(InputBitFile, 8)
'This function works just like inputbit except that it loops through and reads the specified
'number of bits and puts them into a temporary holder
  Dim holder As Long
  Dim longmask As Long
  longmask = 2 ^ (count - 1)
  Do While (longmask <> 0)
    If bitfilename.mask = 128 Then
      Get bitfilename.FileNum, , bitfilename.holder
    End If
    If (bitfilename.holder And bitfilename.mask) <> 0 Then
      holder = holder Or longmask
    End If
    bitfilename.mask = bitfilename.mask \ 2
    longmask = longmask \ 2
    If bitfilename.mask = 0 Then
      bitfilename.mask = 128
    End If
  Loop
  InputBits = holder
End Function
```

