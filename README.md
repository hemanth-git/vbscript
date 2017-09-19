# VBScript
This is the MicrosoftVBscript repo

##Below is the breif scripting introduction and examples.
This file also contains tricky programs on strings and Excel, txt file handling.

#VBScript Tutorial:
-----------------
variables and its types
constant
data types
conditional construct
looping construct
functions and sub procedures
inbuilt functions
object oriented features
FSO model
Excel object model
ADO model
Regular expression
Exception handling

variable : place holder to store the information required for the script

variable types:
vb script supports both explicit and implicit variable declaration.
VBSript has only one type of data type id variant.





implicit variable declaration :

Option explicit : to avoid variables miss use.
‘ :- is for commenting the single line. Vb script provides only single line comments.
Code:
----------
‘Option explicit  for implicit declaration   
‘Dim a
a=5
Msgbox a 

Code:
--------------
a=5;
msgbox(a); 

Explicit variable declaration :
Dim a
a=5
Msgbox(a)


vbscript doesnot hold the declaration and assignment is same line.
variable name must start with alphabets must not exceed more than 225 and can have '_' 
scalar variable: variable that holds single value.

code:
-------------

dim a= 5 'gives error
msgbox(a) 
 
	
Array variables: that can hold multiple values of different.




declaring the constant
--------------------------------

use const to declare the constant.

code:
---------
const b=2


vbscript has only data type is variant
size of variant is 16 bytes.
subtypes of variant are:
-------------------
integer 
boolean
string : with double quotes (")
double
date : with hash (#)
null
empty
object


Scope of variable:
----------------
1.script level :- will be for the entire script, memory allocated for the variable will be deallocated once the script execution is completed.
2.procedure level :- scope of the variable will be only with in the procedure/function in which it is declared any memory will be deallocated onve the procedure is executed.

dim 
public 
private

code:
----------------------------
option explicit
dim a,b,c,d,e,f
a=5
b=6
c=#23-2-2016#
d="hai hemanth"
e=true
msgbox(a)
msgbox(b)
msgbox(c)
msgbox(d)

msgbox(typename(a))
msgbox(typename(b))
msgbox(typename(c))
msgbox(typename(d))
msgbox(typename(e))
msgbox(typename(f))
----------------------------------

runtime input

inputbox("")
return type will be always string.

code:
--------------------------
option explicit

dim a
a= inputbox("enter a value ")
msgbox(a)
msgbox(typename(a))
-----------------------------

operators:
-----------
arithmetic operators
logical operators
relational operators(<,>,<=,>=,=,<>)
concatination operators (&)


"=" is used for both assignment and comparision.


conditional constructs:
---------------------------
block of code written to perform the task based on the conditional
if
if else
elseif
select case
simple if:

	if(condition) then

	block of code

	end if	

type conversion
-----------------
cint(inputbox(""))

looping constructs
----------------------
for loop -- true loop
	
	for i=1 to 5
	
	for var=min to max
	next
	code:
	-----------
	dim a
	for a=1 to 5
		block of code
	next
	
	code:
	----------
	dim a,str
	for a=1 to 5 step 2
		str=str&"hemanth "&a &vbcr
		
	next
	msgbox(str)
	
	code:
	---------
	
	dim a,str
	for a=5 to 1 step -2
		str=str&"hemanth "&a &vbcr
		
	next
	msgbox(str)
	-------------
	
	to break out of the loop use exit
	code:
	-------------
	dim a,str
	for a = 5 to 1 step -1
		
		if(a=3) then
		exit for
		end if
		str=str&"hemanth "&a &vbcr
	next
	msgbox(str)
	
	
for each loop -- true loop
while wend loop -- true loop
do while loop -- true loop
do until loop -- false loop

code for while and do while and do until
----------------------------------------------
	option explicit
	dim a,str
	a=1
	while(a<=5)
		str=str&"hemanth "&a &vbcr
		a=a+1
		wend
	str=str&"this do while iteration" &vbcr

	a=1
	do 
		
		
		str=str&"hemanth "&a &vbcr
		a=a+1
	loop while(a<=5)

	a=1
	do 
		if(a>3) then
			exit do
		end if
		str=str&"hemanth "&a &vbcr
		a=a+1
	loop while(a<=5)
	str=str&"this do until iteration" &vbcr

	a=9
	do 
		str=str&"hemanth "&a &vbcr
		a=a-1
	loop until(a<=5)
	msgbox(str)
-----------------------------------------------------

array hold multiple values of different types
static size is fixed
dynamic size of array can be changed

	dim a(3),i
	a(0)=3
	a(1)=3.3
	a(2)=true
	a(3)="aa"
	
	
	for i=0 to 3
	msgbox a(i)
	next
	redim a(4)


	
for each loop:
----------------------
	no of elements in 
	for each e in a1
	msgbox e


programs:
--------------
/************largest of the three***********/
dim a,b,c

a=cint(inputbox("enter the value of a"))
b=cint(inputbox("enter the value of b"))
c=cint(inputbox("enter the value of c"))

if(a>b and a>c)then
	msgbox(a&" : is the greatest")
elseif(b>a and b>c)then
	msgbox(b&" : is the greatest")
else
	msgbox(c&" : is the greatest")
end if
/************** dynamic array ******************/

option explicit

dim request,dyarray,text,e,str
dyarray=array()
request=inputbox("do you want to enter a value enter 'y' ")
while(request="y")
	redim preserve dyarray(ubound(dyarray) + 1)
	text=inputbox("enter your input ")
	dyarray(ubound(dyarray))=text
	request=inputbox("do you want to enter a value enter 'y' ")
wend
for each e in dyarray
	str=str&", "&e
next
msgbox "your values in array are: "&str

/***************** ATM ***************************/

dim pin,n,reqpin
pin="1234"
n=3

do
	reqpin=inputbox("enter you pin")
	n=n-1
	if(pin=reqpin) then
		msgbox("you have successfully entered correct pin")
		exit do
	elseif(n=0) then
		msgbox("your card is blocked")
	else
		msgbox("wrong pin you have "&n &" chances")
	end if
loop while(n>0)

/*******************************************************/
FUNCTIONS

functions: block of code written to perform the task.

functions or sub procedure is used inorder to make the code reusable.

function function_name()
	block of code
end function
sub procedure_name()
	block of code
end sub

call by value: byval
call by reference : byref
/*************************************************/
code:
----------
option explicit
dim a,b
function addition(byref a,byref b)
	msgbox("this is adding function")
	msgbox("added inside function : "&(a+b))
end function
a=cint(inputbox("enter the value of a: "))
b=cint(inputbox("enter the value of b: "))

function display(byval a)
	a=a+1
	msgbox(" after increament of a : "&a)

end function

call addition(a,b)
call display(a)
msgbox "out the function : "&a

------------------------------

function with return statement is called :procedure
use the same function_name as the return variable
code:-
----------
option explicit
dim a,b,c
function multiplication(byval a,byval b)
	multiplication=(a*b)
end function

function takeinputs(byref a,byref b)
	a=cint(inputbox("enter the a value: "))
	b=cint(inputbox("enter the b value: "))
end function

call takeinputs(a,b)
c= multiplication(a,b)
msgbox c
-------------------

classes in vbscript:

class: its is blue print for an object. it is an imaginary. class can have one or more objects.

object: its is an instance of the class.

method:
code:-
---------

-----------------------

FSO -filesystemobject

class filesystemobject
	method 	driveexist()
code:-
---------
dim fso,d
set fso = createobject("scripting.filesystemobject")
msgbox fso.driveexists("d:")

if(fso.driveexists("c:")) then
set d=fso.getdrive("c:")
msgbox d.availablespace
msgbox d.totalsize
end if


code:-
----------------
dim fso,d,f,f1,str
set fso = createobject("scripting.filesystemobject")
msgbox fso.driveexists("d:")
dim foldername
foldername="c:/testing"
if(fso.driveexists("c:")) then
	if(fso.folderexists(foldername)) then
		msgbox "folder present"
		set f1=fso.getfolder(foldername)
		str=str&" date created "&f1.datecreated&vbcr
		str=str&"date last modified "&f1.datelastmodified&vbcr
		str=str&"date last accessed"&f1.datelastaccessed&vbcr
		str=str&"path "&f1.path&vbcr
		msgbox str
		
	else 
		msgbox "not present"
		set f=fso.createfolder(foldername)
		msgbox "folder created"
	end if
	
	set d=fso.getdrive("c:")
	'msgbox d.availablespace
	'msgbox d.totalsize
	
end if

------------------------
code:--
----------------------------
dim fso,d,f,f1,str
set fso = createobject("scripting.filesystemobject")
msgbox fso.driveexists("d:")
dim foldername
foldername="c:/manual"
if(fso.driveexists("c:")) then
	if(fso.folderexists(foldername)) then
		msgbox "folder present"
		set f1=fso.getfolder(foldername)
		str=str&" date created "&f1.datecreated&vbcr
		str=str&"date last modified "&f1.datelastmodified&vbcr
		str=str&"date last accessed"&f1.datelastaccessed&vbcr
		str=str&"path "&f1.path&vbcr
		msgbox str
		fso.movefolder "c:/manual","c:/testing/"
		f1.move "c:/vbscript/"
		
	else 
		msgbox "not present"
		set f=fso.createfolder(foldername)
		msgbox "folder created"
	end if
	
	set d=fso.getdrive("c:")
	'msgbox d.availablespace
	'msgbox d.totalsize
	
end if

--------------------------------
file operations and methods 
code:-
-----------
dim fso,f
set fso = createobject("scripting.filesystemobject")

dim foldername
filename="c:/testing/manual/new.text"
if(fso.fileexists(filename))then 
	msgbox "exists"
else
	msgbox "no not there"
	set f=fso.createtextfile(filename)
	msgbox typename(f)
	msgbox "file created"
end if
----------------------------------
file writting and append
1-read
2-write
8-append

createfile 
opentextfile
deletefile
copyfile
movefile
getfile
fileexists

path 
parentfolder

-------------------------
code:
--------------------------
dim fso,f,opentxtfile
set fso = createobject("scripting.filesystemobject")
dim foldername
filename="c:/testing/manual/new.txt"
if(fso.fileexists(filename))then 
	msgbox "exists"
	set opentxtfile = fso.opentextfile(filename,1)
	msgbox opentxtfile.readline()
	do
	msgbox opentxtfile.readline()
	loop until(opentxtfile.atendofstream)
else
	msgbox "no not there"
	set f=fso.createtextfile(filename)
	msgbox typename(f)
	msgbox "file created"
end if
---------------------------------------

collections:
-----------
drives:
no methods
properties:
	count
	item

sub folders
methods:
add
properties:
	count
	item

files
properties:
	count
	item

code for getting the drives:
----------------------------
set fso=createobject("scripting.filesystemobject")

set ds= fso.drives
msgbox ds.count
msgbox ds.item("c:")

for each items in ds
	msgbox items
next
---------------------------
sub folders operations:

code:
--------
set fso=createobject("scripting.filesystemobject")

set ds= fso.drives
'msgbox ds.count
'msgbox ds.item("c:")

'for each items in ds
'	msgbox items
'next

set f= fso.getfolder("c:/testing")
set sf=f.subfolders
msgbox sf.count
sf.add("automation")

for each subfol in sf
	msgbox subfol
next
	
----------
operations on files

code:-
------
set f= fso.getfolder("c:/testing/manual")
set sf=f.files
msgbox sf.count


for each subfol in sf
	msgbox subfol
next
	


string inbuilt functions
-------------
left -(string,length)
right-(string,length)
mid - (string,index,length)
join-()
split-(string,delimiter) return type is array
strreverse-()
len-()
ucase-()
lcase -()
replace -(string,"find substring","replace with substring")
ltrim -(string)
rtrim -(string)
trim -()
instr -()


