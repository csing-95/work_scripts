### Makes property blank
```
Replace(Document.Title4, Document.Title4, "") - makes property blank
```

### Finds anything after last "("
```
Trim(Left(Document.Document_Number, InStrRev(Document.Document_Number, "(") - 1)) 
```

### Finds anything with (xxx) and removes
```
Trim(Replace(Document.Document_Number, " (MPC)", " ")) - finds anything with (xxx) and removes 
```