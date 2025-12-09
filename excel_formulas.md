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
=XLOOKUP(A2, A:A, AO:AO, "", 0, -1)
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
for dates formatted yyyy/mm/dd:
=IFERROR(if(AR2="", "", LET(date,AQ8,TEXT(DATE(LEFT(date,4),MID(date,6,2),RIGHT(date,2)),"dd/mm/yyyy"))),"review")
=IFERROR(LET(date,ac2,TEXT(DATE(LEFT(date,4),MID(date,6,2),RIGHT(date,2)),"dd/mm/yyyy")),"")

for dates formatted  xx:
=IFERROR(IF(AR2="", "",LET(date,AR2,TEXT(DATE(LEFT(date,4),MID(date,5,2),RIGHT(date,2)),"dd/mm/yyyy"))),"review")

for dates formatted: mm/dd/yy
=IFERROR(LET(
    date, AF73,
    TEXT(DATE(2000 + RIGHT(date,2), LEFT(date,2), MID(date,4,2)), "dd/mm/yyyy")
), "review")

```

### Turns letter into a decimal, good for reordering by revision number if it goes into double digits (rev no. 10 can appear before 1)
```
=CODE(RIGHT(A2,1))/100
```

### For making digit 1 -> 01
```
=IF(LEN([@[Sheet Number]])=1,"0"&[@[Sheet Number]],[@[Sheet Number]]) &""
```

### Pulling latest revision from stack in Masters
BY = Title 4 (Document Number from Projects), I = Revision number, H = Title 4 (Document Number from Masters)
```
=LET(
 stack, BY6,
 revs, FILTER('[007 - Wascana Masters Loadsheet.xlsx]Documents'!I$2:I$255, '[007 - Wascana Masters Loadsheet.xlsx]Documents'!H$2:H$255=stack),
 INDEX(revs, COUNTA(revs))
)
```

### Identifying XREFS from title and document name/number
```
=IF(
  ISNUMBER(SEARCH("3D MODEL", V2)),
  "XREF 3D MODEL",
  IF(
    OR(
      ISNUMBER(SEARCH("xref", V2)),
      ISNUMBER(SEARCH("x ref", V2)),
      ISNUMBER(SEARCH("x-ref", V2))
    ),
    "XREF TITLE",
    IF(
      AND(
        SUMPRODUCT(LEN(H2)-LEN(SUBSTITUTE(UPPER(H2),CHAR(ROW(INDIRECT("65:90"))),""))) > 1,
        ISERROR(SEARCH("95019", H2))
      ),
      "XREF PROP",
      ""
    )
  )
)
```

### Formatting Revision numbers with legacy revs numbers(0.0)
```
=IFERROR(IF(F2="False", CONCAT(I2, "(", K2, ")"),I2),"")
```

### Identify dupe stacks with hybrid issues (will slow down if large dataset)
```
=IF(SUMPRODUCT(($A$2:$A$9999=A9)*(COUNTIFS($G$2:$G$9999,$G$2:$G$9999)>1))>0,
    "Dupe stack",
    ""
)

-------------

Step 1: (g=document number+rev no.):
=COUNTIF($g:$g, g2) > 1

Step 2 (j=column used for step 1): 
=COUNTIFS($A:$A, a2, $J:$J, TRUE) > 0
```