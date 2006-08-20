
-----------------------
COMPILE FROM THE SOURCE
-----------------------
(Internal remark: The code in this folder is just a copy from _Dev)


FLEXCEL LIBRARY

- xlReadWrite has been compiled with inst\binary\xlsReadWrite.dll
- To fully compile the code you need Flexcel, a third party library.
  See: www.tmssoftware.com/flexcel.htm
- As this is a non-free library, my code contains an explicit 
  exception to allow linking it to Flexcel

- In Flexcel derive TFlexCelImport and TXLSAdapter from TObject 
  (instead of TComponent). Just change this and also adapt the 
  correspondending Create methods
- Flexcel has quite a lot of different folders. I just copied 
  all .pas, .inc and .res files in a single folder
- In Delphi set a path to this folder


DELPHI

- I used Delphi 2006. For the headerfile-conversion-project I also 
  tried out D6, which worked flawlessly. Versions below D6 
  won't work because of the varargs directive.
- In Run->Parameters set the host application 
  (e.g. C:\Programme\R\R-2.3.1\bin\Rgui.exe). 
- Breakpoints etc. just work fine. 


DOCUMENTATION

- You can use the Writing R Extensions manual as the functions 
  resemble closely their C counterparts. 
[- Maybe it is a good idea anyway to start with the demos in the 
  headerfile-conversion-project. 
  You can download this fully GPL'ed project from:
  http://treetron.googlepages.com (RHeaders4Delphi.zip)]
  