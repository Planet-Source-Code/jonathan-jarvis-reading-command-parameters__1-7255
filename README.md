<div align="center">

## Reading Command Parameters


</div>

### Description

Have you ever wanted to click a file and open it? If your program is associated with a specified file type, or you select file(s) and drag them onto the exe icon of your app then this code is for you. It reads the parameters when the form is loaded using by getting the "Command".

This code can get each file dragged or clicked on. If you just use the command statement you will get something like this "c:\george.bmp c:\ben.bmp f:\mydoc.doc". This code seperates each file and uses another sub to open it.
 
### More Info
 
If you want to use this code so far you will need a listbox with a name of "List1"

This uses some some simple string work and a loop statement.

THis code returns seperated command parameters

None as of yet


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan Jarvis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-jarvis.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-jarvis-reading-command-parameters__1-7255/archive/master.zip)





### Source Code

```
'Code created by Jonathan Jarvis - Jman
'Email: roboman1@email.com
'as of now this will require a listbox with a name of "List1"
Private Sub getfiletoopen(filename As String)
List1.AddItem filename
End Sub
Private Sub Form_Load()
'create variables
Dim howlong, n As Integer, c As String
'give variables values
c = Command
n = 1
For howlong = Len(Command) To 1 Step -1 ' start loop statement
If Mid(c, n, 1) = " " Then 'check to see if It should seperate commands
getfiletoopen Mid(c, 1, n - 1) 'pick out the command from line only Mid(c, 1, n - 1) is the command file
c = Right(c, Len(c) - n) 'change command and get rid of last handled file
n = 0 'reset letter to 0
End If
n = n + 1 'increment to next letter
Next howlong 'go on to next letter
'takes care of last command line or 1st one if only one file is to be opened
If c <> "" Then ' checks to see if there is a 1st or last command
getfiletoopen c ' you can change this to load your file or command. c is the command parameter of the last file
End If
End Sub
```

