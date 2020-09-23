Download the Socket.dll reference here:

http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=39858&lngWId=1


You can add it in 2 ways.

1) add the dll project to the server project and save it as a group project. Don't just add the classes and modules, add it as a seperate project. Re-reference the dll, and away you go.

2) Compile the dll and then open the server project. Unreference it and press OK. Cliek the references option again, and browse for the compiled dll, add it and press OK.


Windows XP users, I included the manifest file for the project so if you compile the server, it should have the WindowsXP look as seen in the screenshot.