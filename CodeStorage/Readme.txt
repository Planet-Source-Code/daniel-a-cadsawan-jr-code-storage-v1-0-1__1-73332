'----------------------------APPLICATION----------------------------------
'App name: CS (Code Storage)
'App desc: Utility to store codes, types from different language,
'   or even notes from batch, pc registry, access or excel programming, etc.
'App author: Daniel A. Cadsawan Jr.
'Author info: xavierjohn22@yahoo.com, +639209190715
'http:\\cad3dmdd.ucoz.com
'http:\\www.cad3dmdd.com

'--------------------------ACKNOWLEDGEMENT---------------------------------
'Special thanks to the following, I modified most on these owner's code
'LGS Static's Code on colorcode, wordlist, find/search codes
'Static, I redo the database but made the samples from Codebank appear herein
'Napparan Philip's style on database connecting and handling
'User Controls acknowledgement as follows:
'McToolBar 2.3 by Jim Jose
'La Volpe Buttons vH.1

'-------------------------------HISTORY------------------------------------
'CS (Code Storage) was inspired by my Windows 7 laptop where i experienced
'a lot of error and also registering ocx from CodeBank
'I was using CodeBank Version: 4.0.0, Geoff Goldsmith -aka [LGS]Static
'Co-Author is me, Daniel A. Cadsawan Jr. -aka xavierjohn22 only
'bumping it several revisions to get away from the pc registry settings
'I needed to remake this for Windows 7, got rid of the OCX dependency,
'and completely redo pretty much everything, only the connection is active
'For this i connect and get records set the recordset to nothing
'in the future i will cut the connection to the database as well

'---------NOTE ONLY IMPT LIST WHERE THIS PROJECT IS REFERENCED TO----------
'msjro.dll, msbind.dll, msado25.tlb