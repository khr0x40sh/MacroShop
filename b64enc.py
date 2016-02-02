#!/bin/python

import os, sys

#OPEN THE FILE
if os.path.isfile(sys.argv[1]): todo = open(sys.argv[1], 'rb').read()
else: sys.exit(0)

#ENCODE THE FILE
print "[+] Encoding %d bytes" % (len(todo), )
b64 = todo.encode("base64")

#WRITE THE OUTPUT
print "[+] Encoded data is %d bytes" % (len(b64), )
f = open("base64_output.txt", 'w')
f.write(b64)
f.close()
print "[+] Done!"

#WRITE OUTOUT TO VB FRIENDLY FORMAT
print "[+] Writing to base64_output.vb"
vb_in = open("base64_output.txt", 'r')
i = 0
str = 'Dim var1\n'
for line in vb_in:
	line = line.strip("\n")
	if i > 0:	
		str = str + "var1 = var1 & \"" + line + "\"\n"
	else:
		str = str +"var1 = \""+ line+"\"\n"	
	i=1
vb_in.close()
f = open("base64_output.vb", "w")
f.write(str)
f.close()
print "[+] VB file completed!"
