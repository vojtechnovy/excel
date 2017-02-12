# excel

# look-up string of characters in C1 matching text cells A1:A3 and returns value from matching cell in column B
# C1 and A3 = "QWERTY", will return B3
=IFERROR(INDEX($B$1:$B$3,MATCH(1,--NOT(NOT(FIND($A$1:$A$3,C1))),0)),"N/A")
