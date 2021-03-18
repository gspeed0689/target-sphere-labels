# target-sphere-labels
An older python script to create randomized code labels to be wrapped around laser scanning sphere targets for identifying individual locations of the target. 

Terrestrial laser scanning uses spherical targets to help register data.
For individual placement tracking purposes I wanted a script to randomly generate lables to be attached to the tripods the sphere targets are placed on, and thus this script was born. 
In Microsoft Word I created a template document and saved it as an `xml` document file compatible with Word.
This script reads the template file, replaces the text using proper `xml` document object model manipulation with the `xml.dom.minidom` python library.
The script will automatically call Word with the `subprocess` library and also automatically print the randomly generated codes document to the system's default printer. 
