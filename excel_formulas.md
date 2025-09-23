### For sheet number
Extracts the last 2 digits from the document name (Double check the output!)
```
=LET(
  txt, [@[Document Name]],
  base, LEFT(txt, FIND(".", txt&".")-1),
  hyps, LEN(base)-LEN(SUBSTITUTE(base,"-","")),
  start, FIND("~", SUBSTITUTE(base,"-","~",hyps)),
  tail, MID(base, start+1, 2),
  rest, MID(base, start+3, 1),
  IF(AND(OR(hyps=3, hyps=4, hyps=5), ISNUMBER(--tail), LEN(tail)=2, rest=""), tail, ""))
```  

### For sheet count
Gets the last sheet number in the stack (Double check the output!)
Column A = Stack ID, AH = Sheet Number
```
=LOOKUP(2,1/($A$2:$A$739=A2),$AH$2:$AH$739)
```

### Check stack id for rendition review
Column A = Stack ID, P = Rendition Path
```
=A2 & ": " & IF(COUNTIFS(A:A,A2,P:P,"*Review*")>0,"Review","No Review")
```

### Formula for stack ID based on Document Number:
This creates stack ID based on the document number
Column E = Document number
```
="Stack_" & TEXT(MATCH(E2,UNIQUE($E$2:$E$31),0),"000")
```

### Date formatting (Double check the date is the correct date and has formatted correctly)
```
=IFERROR(LET(date,AQ8,TEXT(DATE(LEFT(date,4),MID(date,6,2),RIGHT(date,2)),"dd/mm/yyyy")),"")
=IFERROR(LET(date,AD2,TEXT(DATE(LEFT(date,4),MID(date,5,2),RIGHT(date,2)),"dd/mm/yyyy")),"")
```

### Turns letter into a decimal, good for reordering by revision number if it goes into double digits (rev no. 10 can appear before 1)
```
=CODE(RIGHT(A2,1))/100
```