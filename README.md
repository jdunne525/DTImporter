HID Descriptor tool import tool
===========

The version 2.4 HID descriptor tool from usb.org tragically doesn't contain an import .c file functionality.
It only supports its own proprietary .hid file format (which fortunately is fairly simple).
This tool is written in VB6.

This tool imports a .c file containing only the report descriptor and outputs a .hid file containing the descriptors
imported.  Note that DT.exe will choke when trying to parse the descriptor unless actually COPY the .HID file to the
folder where the dt.exe tool is located.

Limitations of the tool:  (I threw this together in 2 hours, so don't expect perfection)
//The tool properly ignores // comment lines and BLANK lines, and now DOES support C style comments /* */
//The tool looks for a comma to determine if the line contains data.
//The tool only supports ONE ITEM PER LINE!!  
//So, if you've got 0xC0, 0xC0 on a single line to end 2 collections, the tool will FAIL to import that data.

Updated DT.exe's library files to the latest files from the USB-IF CV test
suite.

DT.exe works with the new dll's and upg's.  
To update HID descriptor tool using USB20CV files, drag in these files from USB20CV and replace them all.

C:\Program Files\USB-IF Test Suite\USB20CV\lib\
- Tparse.dll
- dt.ini
- *.upg

