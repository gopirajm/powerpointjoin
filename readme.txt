PowerPoint Join
Author:  Richard Sugg (richardsugg@gmail.com)

* note that sprintf.js isn't mine.  Read the file to see where it came from

Use
*	To use, run "pptjoin.hta".  Follow the instructions on the screen.
*	To generate a textfile from a directory full of charts, open a DOS
	prompt and run the following command:

	dir /b *.ppt* > my_files.txt

	This will list all ppt and pptx files and store their names in my_files.txt.
	The files will be listed in alphabetical order.  If you want them combined
	in a different order, you must rearrange them in this file.
*	To translate to another language, make a copy of lang.<lang>-<country>.js
	and translate the strings.  Be careful not to modify the "%s" and any
	embedded HTML or javascript in these strings.  You will also need to set the
	application to use the new language file by changing the line in pptjoin.hta 

	<script type="text/javascript" src="lang.en-US.js">

	The src attribute must be your language file.

Prerequisites
*	Text file containing the list of charts to combine must be in the same
	directory as the charts that are being combined.
*	Office 2003/7 (duh)
