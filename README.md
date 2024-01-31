# Excel-Formulas
Some advanced Excel formulas for referance:-



=IF(ISNUMBER(MATCH(C6,$A$3:$A$11246,0)),"YES","NO")     -----> This will search IF the value in c6 exist in any cell starting from A3 to A1126.

=IF(Sheet1!E3="YES",Sheet1!C3,"0")      ---------> This will copy value from Sheet1 C3 IF Sheet1 E3 value is = YES.

=IFERROR(VLOOKUP(A2,Sheet1!$A$3:$B$11245,2,TRUE),0)   ------->This will copy the value from Sheet1 B column IF value in A2 exist in sheet1 A column.
                                                              Note: Value copied from Sheet1 B will in the same row where A2 exist in Sheet1 A

=IFERROR(INDEX(Sheet1!$J$3:$J$18142,MATCH(1,(B2=Sheet1!$H$3:$H$18142)*(C2=Sheet1!$I$3:$I$18142),0)),"0")    -----> this will copy the value from Sheet1 J column, IF
values from Sheet2 B2 is = Sheet1 H Column and Sheet2 C2 is = Sheet1 I column values.
Note: Both matching values for B2 & C2 should exist in same row in sheet1 in column H & I
      Value copied from Sheet1 J will also be from the same row where above match exists.

=IF(AND(D2<=G2,E2>=G2),"AP",IF(OR(D2=0,E2=0),"BLANK","NR"))   -------> This enter the values AP/NR/BLANK in the the cell, based on the value/number comparation
this is trypical If, Else, AndElse condution formula.


=IF(AND(H2="AP",I2="AP"),"T1",IF(AND(H2="NR",I2="NR"),"T2",IF(AND(H2="AP",I2="NR"),"T3",IF(AND(H2="NR",I2="AP"),"T4","Missing data"))))  ----> This is similar to above formula with more condations.

=COUNTIF(Sheet3!$J$2:$J$87289,"T1")  ---->  This normal count operation for the given column, It will enter count of no of occurances of T1 in column J.

=UNIQUE(Sheet3!$C$2:Sheet3!$C$87289)   -----> This will list the all the unique values exist in the Column C.

=COUNTIFS(Sheet3!$C$2:Sheet3!$C$87289,A4,Sheet3!$J$2:Sheet3!$J$87289,"T1")  ----> This will enter the count of occrance where A4 value matchs with Sheet3 column3 and value T1 matchs from sheet3 column J.   In both matching conditions the values must be in same row.


