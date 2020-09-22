What does Oil do?
Uses a complied executable to create an update installation package
Finds the executable's File name and version
Allows for entry of new files required for the update.
Allows for creation of new folders/directories.
Allows for new files to be .zip compressed
Allows for update to the uninstall log - ST6UNST.LOG, for VB's P&D wirard
Creates an install ini file that the Fuel install reads.
Copies any new files and the Fuel installer to the package folder.

What does the Fuel installer do?
Reads the install.oil (ini file)
Finds the Application to update's path
Finds the Windows System Dir.
Determines what files to install and where they go.
Installs the files, downloads then if needed. 
Registers any System files (ocx,dll)
Updates the uninstall log - ST6UNST.LOG,  if selected to.

Make sure the ZipFunctions.class is installed 
See the readmfirst.txt in the zipclass folder.