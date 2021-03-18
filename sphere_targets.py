from xml.dom import minidom as xdmd
import subprocess, os, random, string, uuid

#Path to your WINWORD.EXE program
word_path = r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
#Path to the template xml document file 
template_path = r"C:\cmd\qpf.py\SphereTargetTemplate.xml" 
#Path to a temporary location for writing a new xml file to
temp = r"C:\temp"

def load_xml_dom():
	"""This function loads the template xml file into an xml.dom.minidom object
		The global template_path needs to be set
		
	INPUTS:
		global path for a template word document
	
	RETURNS:
		xml.dom.minidom object of the template document file"""
	a = open(template_path, "r", encoding="utf-8").read()
	b = xdmd.parseString(a)
	del a
	return(b)
	
def generate_random_letters():
	"""generate_random_letters() returns a dictionary of 6 randomly generated 3 letter strings
		The keys for the dictionary match the values that will be replaced in the xml word document
		The random letters are upper case ASCII letters
		
	INPUTS:
		None
	
	RETURNS:
		dictionary of standard keys and randomly generated values"""
	d = {"ONE":"", "TWO":"", "THR":"", "FUR":"", "FIV":"", "SIX":""}
	for i in d.keys():
		for j in range(3):
			d[i] += random.choice(list(string.ascii_uppercase))
	return(d)
	
def replace_letters(xd, d):
	"""replace_letters(xd, d) replaces the default text with the randomly generated text

	INPUTS:
		xd is the xml.dom.minidom object loaded in from load_xml_dom()
		d is the dictionary object created from generate_random_letters()
		
	RETURN:
		xd is returned after being modified with the values from d"""
	for i in xd.getElementsByTagName("w:t"):
		r = i.firstChild.nodeValue
		nv = xd.createTextNode(d[r])
		i.replaceChild(nv, i.firstChild)
	return(xd)

def write_xml(xd):
	"""writes the xml.dom.minidom object to a temporary file. The temporary file name is
		a uuid 4 hex string with a .xml extension
	
	INPUTS:
		xd is the xml.dom.minidom object from replace_letters(xd, d)
		global temp folder path 
	
	RETURN:
		path to the temporary word xml document location"""
	u = uuid.uuid4().hex
	p = temp + os.sep + u + ".xml"
	f = open(p, "wb")
	f.write(xd.toprettyxml().encode("utf-8"))
	f.close()
	return(p)
	
def print_document(u):
	"""print_document(u) uses command line arguments to call the word executable and print
		the temporary file to a system printer. 
	
	INPUTS:
		u is the path to the temporary word xml document
		global word_path is the path to the WINWORD.EXE executable
	
	RETURNS:
		None"""
	subprocess.Popen([word_path, u, "/mFilePrintDefault", "/mFileExit"]).communicate()
			
def main():
	"""main() runs the different modules in their needed order"""
	xd = load_xml_dom()
	d = generate_random_letters()
	xd = replace_letters(xd, d)
	u = write_xml(xd)
	print_document(u)
	os.remove(u)

main()