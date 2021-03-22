# target-sphere-labels
An older python script to create randomized code labels to be wrapped around laser scanning sphere targets for identifying individual locations of the target. 

Terrestrial laser scanning uses spherical targets to help register data.
For individual placement tracking purposes I wanted a script to randomly generate lables to be attached to the tripods the sphere targets are placed on, and thus this script was born. 
In Microsoft Word I created a template document and saved it as an `xml` document file compatible with Word.
This script reads the template file, replaces the text using proper `xml` document object model manipulation with the `xml.dom.minidom` python library.
The script will automatically call Word with the `subprocess` library and also automatically print the randomly generated codes document to the system's default printer. 

**Usage**

Place all of the provided files in the same directory, and edit the python file with a path for your Word executable, the path to the template `SphereTargetTemplate.xml` file, and a path for a temporary directory. The three variables to change are all at the top of the python file after the imports. You should then be able to run `python sphere_targets.py` as many times as you like and it will print six 3 letter code labels for attaching as labels to scanning sphere targets. 
