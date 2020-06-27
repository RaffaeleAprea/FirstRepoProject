# Excel finder by cell value
This project will consist of an utility meant to look for all Excel files that
hold inside a string passed by the user. All files analysed will be those
inside node and leaft sub-Directories of the Folder where the script is issued.
So the searching is recursive.

This script helped me a lot to check some stuff litterly hided inside a huge set
of excel configuration files of a CAD software.

The script uses COM automation, infact two Excel instances are created:
1. The first one is invisible and opens one by one all the excel files found
and searches the input string in the cells of each one.
2. The second one is visible and at first shows all excel files found, and then
shows a report which gets populated at run-time as it shows the status of the
searching of the string inside all excel files listed in a column.
