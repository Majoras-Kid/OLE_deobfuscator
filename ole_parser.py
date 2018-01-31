makro_text = open("makro2.txt","r")
parsed_python = open("parsed.py","w")
makro_function_text = open("makro1.txt","r")

global variables
global variables_definition
global functions
global call_functions
global function_names
mid_function = "def mid_function(string, index,length):\n"
mid_function += "\tstring = string[index-1:index-1+length]\n"
mid_function += "\treturn int(string)\n"

def mid_func(string,index,length):
	string = string[index-1:index-1+length]
	return int(string)

def parse_line():
	counter = -1	
	variables = ""
	variables_definition = ""
	functions = ""
	function_names = ""
	functions += mid_function
	functions +="\n"
	call_functions = ""	
	#for line in makro_text:
	while True:
		counter = counter + 1

		line = makro_text.readline()
		if not line:
			break

		#print "Current line: %s" % line

		if "Attribute" in line:
			print "Attribute Line"
			continue
			
		#String definition
		#Public vjspg As String	
		if "Public" in line:
			line = line.replace('Public ', 'global ')
			line = line.replace('As String\n','')
			variables_definition += "%s = \"\"\n" %line
			call_functions += "print %s\n" % line.replace('global','')
			#line += "=None\n"
			
			variables += line
                        variables += "\n"
                        continue


		#function definition
		if "Sub " in line:
			#for loop until 
			print "Function definition: %s" % line
			line = line.replace('Sub ','')
			#call_functions += "print %s" % line
			#function_names += line
			line = line.replace('()\n','')
			functions += "def %s():\n" %line
			while True:
				next_line = "\t" +makro_text.readline()

				if "Array(" in next_line:
					next_line = next_line.replace('Array(','[')
					next_line = next_line[:-2]
					next_line += "]"
					#print "Next_line: %s" % next_line
				else:
					if "End Sub" not in next_line:
						variable_name = next_line.split(' ')[0]
						#next_line = next_line.replace('\n\t','\n\trturn ')
						#print "Return value: %s" % next_line
						next_line = next_line.replace('(','[')
						next_line = next_line.replace(')',']')
						next_line = next_line.replace('= ','=\"\".join([')
	                                        next_line = next_line.replace('\n','])\n')
						next_line = next_line.replace('&',',')
						next_line = next_line[1:]
					
						variable_name = "global %s\n"% variable_name[1:]
						variable_name = "\t" + variable_name 
					
						print "Variable: \n%s"% variable_name			
						next_line = "\t" + next_line
						next_line = variable_name + next_line
						print "Variable adding: %s" % next_line
				line += next_line
				next_line = next_line.replace('VBA.Mid$', 'mid_function')				
				
				if "End Sub" in next_line:
					functions +="\n"
					break
				next_line += "\n"
				functions += next_line
				#functions +="\n"
				#print "Funktion is: %s" % line
			#print "Function string %s" % line
			continue

	while True:
		line = makro_function_text.readline()
		
		if not line:
			break
		if "qtkyh." in line:
			print "Line: %s"%line
			line = line.replace('qtkyh.','')
			line = line.replace('\n','()\n')
			function_names+=line
		
	#print "Variables: %s" % variables
	#print "Functions: %s" % functions				
	parsed_python.write("#Variables\n")
	parsed_python.write(variables)
	parsed_python.write("\n")
	#parsed_python.write("#Defining Variables\n")
	#parsed_python.write(variables_definition.replace('global ',''))
	parsed_python.write("#Functions\n")
	parsed_python.write(functions)
	parsed_python.write("#Calling functions\n")
	parsed_python.write(function_names)
	#parsed_python.write(call_functions)
	parsed_python.write(call_functions.replace('print ',' ').replace('\n',''))

parse_line()

#print "Testing mid_func"
#char = (mid_func("98f6lb050", 8, 2))
#print "%s" % char
#print "chr(char): %s" % chr(char)

