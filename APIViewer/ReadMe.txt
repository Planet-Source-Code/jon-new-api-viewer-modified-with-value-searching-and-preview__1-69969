NEW API VIEWER 1.0.0
Copyright 2007 © David Ross Goben
-----------------------------------------------------------------------------
This Add-In is a workable replacement for both the stand-alone and Add-In API Viewer that comes with VB6, and is loaded with features that most everyone had wished the original versions had.

CONTENTS:
--INTRODUCTION
--RECOMPILING THE ADD-IN VERSION
--INSTALLING THE NEW API VIEWER
--NEW FEATURES BEYOND THE STANDARD VB6 API VIEWER
--IMPLEMENTATION NOTES
--ADD-IN DEVELOPMENT NOTES

------------
INTRODUCTION
------------
The New API Viewer comes in 2 flavors: As a stand-alone executable, and as an Add-In to the Visual Basic 6 Integrated Development Environment (IDE), accessible from the Add-Ins menu. The main difference between the two is that the Add-In version features one additional button, Insert into VB Code, which allows the selection API data to be inserted directly to the VB source code, if a code window is currently open. Regardless, it will also place a copy of the selected API data in the clipboard, as though the Copy to Clipboard button had also be pressed.

------------------------------
RECOMPILING THE ADD-IN VERSION
------------------------------
If you have downloaded updated source code for the New API Viewer, you will need to recompile the source code (obviously). What might not be so obvious to newer users is that the ActiveX DLL will need to be unloaded from VB before recompiling it. But this is a very easy thing to do, so there is nothing to worry about. Actually, the steps required for doing this in general are noted below in "ADD-IN DEVELOPMENT NOTES", but here are the quick steps for the people new to ADD-IN compiling. Mind you, these steps are only required if you already have a previous version of the New API Viewer Add-in built:

1) Load the new source code into VB, as usual.
2) From the Add-Ins menu, select "Add-In Manager..." to bring up the Add-In Manager dialog. Select the "New API Viewer" entry in the list. Make sure that all checkboxes are unchecked in the "Load Behavior" frame. Click OK to exit the Add-In Manager.
3) Recompile the DLL (we had to ensure that the Add-In was unloaded, per step 2, otherwise we could not recompile the DLL because it would be in use).
4) After a successful recompile, bring up the "Add-In Manager" again, select the "New API Viewer Entry", and ensure that the "Loaded/Unloaded" checkbox in the "Load Behavior" frame is checked (this step would be automatic if you left VB and then came back in).

You can now use the updated Add-In version of the New API Viewer.

-----------------------------
INSTALLING THE NEW API VIEWER
-----------------------------
If you have been provided with the APIViewer.Exe and APIViewer.DLL files from someone who has compiled them under VB6, or you simply want to install them on another system without having to recompile them on that system as well, do the following:

1) Create a new folder, or choose a folder that you would place common executables and DLL files. Many people prefer the System32 folder under the Windows folder, but this is a place where way too many other files reside, and it is difficult to keep track of the files that you will personally maintain. This choice is, as always, yours. Regardless, choose a folder that you will not later delete without considering that it is important. Place the EXE and the DLL versions of the New API Viewer into that location.

2) Create a Desktop shortcut for the APIViewer.exe file, and then rename and move the shortcut to wherever it will be convenient for you, such as a desktop tools folder, the Quick Launch bar, or any other convenient location.

3) Register the APIViwer.DLL file. Although this can be a difficult task. There are two methods that make such tasks easy. The first is to install the RegUnreg.reg file, which came with the APIViewer source code. This will add two important options to your context menu when you right-click DLL and OCX files within Windows Explorer: entries titled Register and Unregister. These allow you the convenience of registering and unregistering Component Object Model (COM) DLL's and OCX files with simple and quick button clicks. See the RegUnreg.Txt file for additional details. The second method is to create a shortcut to the system's RegSvr32.exe file (found in your system32 folder under your windows directory) and place the shortcut into your SENDTO folder (to bring this folder up, go to Start / Run, and enter SENDTO to open a browser to this location). Placing a shortcut to Regsvr32.exe in this folder will allow you to rightclick a DLL or OCX file, select Send To, and then RegSvr32.exe to register a file.

4) Although this step is normally not necessary because the New API Viewer is configured to tell VB that it is in fact a VB6 component, the following is provided for informational purposes only, in case something in the system happened that prevented VB6 from automatically incorporating the New API Viewer into its Add-Ins menu:, Bring up Visual Basic, and select the Add-In Manager from the Add-Ins menu (if you see "New API Viewer" in the Add-Ins menu, then this is not necessary). Select the New API Viewer entry, and ensure that the Loaded/Unloaded and the Load on Startup checkboxes are checked in the Load Behavior frame. Select OK. You can now activate the Add-In version of the New API Viewer by selecting its entry from the Add-Ins menu. 

You are now ready to use the New API Viewer.


-----------------------------------------------
NEW FEATURES BEYOND THE STANDARD VB6 API VIEWER
-----------------------------------------------
This Add-In is a workable replacement for the API Viewer that comes with VB6 loaded with the features that most everyone wants.

* It gives you the ability to declare dynamically constants as Long in the selected local copy of a constant. This allows the user to keep program speed optimal by not slowing down for variant conversions. Although constants can be of types other than Long, the VB6 API interface uses only Long Integer values for its constants.

* It gives you the ability to create new constants right within the viewer. Assigned values are expected to be numeric, as is required by the VB6 API interface, but the values can be declared as hexadecimal, octal, or binary. You can also apply + or - offsets. Constants are normally created as all-capitals. No complex checks are performed on the value. It simply assumes that you know what you are doing, because such checks can involve complex offsets and naming of other constants. The viewer will also check to ensure that the newly entered constant does not already exist.

* It gives you the ability to create new API Declarations right within the viewer and add them to your API list.

* It gives you the ability to create new User-Defined Types right within the viewer and add them to your API list.

* It gives you the ability to create new Enumerations right within the viewer and add them to your API list.

* You can Delete entries from the API list.

* It gives you the ability to edit Declared Subroutine and Function parameter lists, and apply these changes to new subroutine or function names (for example, saving a  modified version of SendMessage to SendMessageByNum, after changing the lParam as Any to ByVal lParam As Long). The built-in Declaration Editor makes such changes a breeze with just a few clicks of the mouse.

* It automatic checks for new parameter and constant dependencies. If an added declaration or user-defined type or constant requires another user-defined type or constant not included in the selection list, you can view the requested types in a dialog and select them or reject them for inclusion in your selection list. This can make resolving declaration headaches such as with the complicated AccessCheck Function, which requires the additional inclusion of the GENERIC_MAPPING, PRIVILAGE_SET, and SECURITY_DESCRIPTOR types. These additional types in turn require the ACL and LUID_AND_ATTRIBUTES types. These newer types in turn also require the LUID type. The New API Viewer makes farming these additional types a breeze with a few quick clicks of the mouse.

* Additions created within the New API Viewer can be optionally saved for later re-used in the API data file. New entries are appended to the API file with a date- and time-stamp marker.

* It immediate displays updates when you toggle between Private and Public declaration options, define parameters as arrays or fixed-length strings, change a return type, or change the referencing verb (ByRef/ByVal or None).

* It includes a copy of and richly expanded freeware API32.TXT by Dan Appleman (president of Desaware, Inc.), which he had derived from the original Win32API.TXT file. On top of that, I have also included several thousand new Constants, Declarations, and Type declarations, to include some undocumented hooks and declarations that get around some entries that were thought impossible in VB. This new file should be copied to your Visual Studio Folder (Usually C:\Program Files\Visual Studio 6), in the Common\Tools\Winapi folder).

* Two Versions of the New API Viewer are included. A stand-alone version that compiles to an EXE that can be launched outside of VB (this is useful for inclusion in other language development apps, though declaration syntax may need to be altered), and an Add-In version that compiles to an ActiveX (OLE) DLL and runs from within VB via the Add-Ins menu. Notice that in the stand-alone version, the INSERT INTO VB CODE button is hidden, because it has no direct communication with the VB IDE).
  

---------------------
IMPLEMENTATION NOTES:
---------------------
Although this version of the API Viewer does not create an Access Database from the API text file, due to the extremely fast file-loading technique I employ, I found that such support, though easy to implement, was not necessary. The clock-speed issue that was a problem on older systems, which had made such a conversion advantageous, is now moot and no longer offers any significant advantage. Further, because the viewer works with the original API text files, the ability to dynamically add new constants and create modified subroutine and function declarations allows easy transport of such changes through most mediums to other users via that text file format.


-------------------------
ADD-IN DEVELOPMENT NOTES:
-------------------------
Here are a few things you might try to make building and rebuilding ADD-INs a lot less of a bother.

1)When you build an ADD-IN as a DLL, place the DLL in a place that you might later not destroy when you are doing house-cleaning. If you do need to move the DLL file, be sure to first unregister it. Follow these steps: 1) Unregister it (this is actually optional, as long as you are sure to re-register it after moving, but unregistering it is simply safer, in case the kids interrupt with the latest disaster, distracting you), 2) Move it to its new location, and 3) re-register it. If you are not sure how to register and unregister a file, see the included RegUnRegHELP.txt help file. This will make registering and unregistering register-able DLL and OCX files a breeze with the included RegUnreg.reg file.

2) After you have built your Add-In DLL, simply exit VB and then go back in. This will load your DLL and add it to the Add-Ins Menu. Alternatively, instead of exiting VB, simply select the Add-In Manager from the Add-Ins menu, select the title of you ADD-INs' entry, and ensure that the "Loaded/Unloaded" option is checked (loaded), then hit OK and exit the manager. You can now select your new ADD-IN title from the Add-Ins menu and begin using it.

3) If you find a bug in you ADD-IN and need to build a new one, load you ADD-IN project for editing, then immediately select the Add-In Manager from the Add-Ins folder, select the desired ADD-IN (the project you are working on), and make sure all checkboxes under LOAD BEHAVIOR are unchecked. You can now fix or tweak it. Actually, this precaution is only needed for building the new DLL, but it is a good practice to work on your ADD-IN code while the ADD-IN DLL is unloaded.

4) Per Point 3 above, it is a good idea to first build a stand-alone version of the application for testing and primary debugging (I first create a Stand-Alone version, and then created the ADD-IN only after I have everything else working). Once you have completed stand-alone testing, Create an Add-In Project with either a slightly different name or place it in a sub-folder named something like, Add-In Version. Be sure to load all forms and modules used by the stand-alone version (I cheat and edit the VBP files in a text editor and copy most of the common data. Also, change the add-in project's startup form to your main form, not the provided frmAddIn. Finally, add the following lines to the top of your main (startup) form while in the Project:

  #If ISADDIN = 1 Then
    Public VBInstance As VBIDE.VBE
    Public Connect As Connect
  #End If

Then, at any point in your main form where you have an Unload Me, replace it with the following lines:

  #If ISADDIN = 1 Then
    Connect.Hide
  #End If
  Unload Me

Also, edit the code for the Connect object (in the Designers folder, when creating your ADD-IN version), and replace the two instances of "frmAddin" with the name of your main form (the first for the main form declaration, and the second for instantiation). After doing all this, you can remove the frmAddIn form from your project.

Finally, in your project properties under the MAKE tab, in the conditional compilation line, add "ISADDIN=1" (by the way, change this to "ISADDIN=0" for you stand-alone project file; this way you can use the very same code and forms, as I have done with the New API Viewer application. Search the Stand-Alone or Add-In version in the current project for instances of ADDIN, and see how I managed blending both the stand-alone and ADD-IN versions together).

5) You might notice that you have to open the CONNECT object in the Designers VB folder (from inside the VB environment with your ADD-IN loaded) and set the Add-In Display name to your project (and also set Application Version, Initial Load Behavior, and Description). But there is much more. Open the code view for the Connect object. Search the code for "My Addin", with the quotes [the function will be: AddToAddInCommandBar("My AddIn")]. Replace the quoted text with the name you want to be displayed for your ADD-IN in the Add-Ins dropdown menu. You should also edit your ADD-IN's Project Properties. You will notice that the description on the General Tab contains the text "Sample AddIn Project". You should change the text to a better description for your ADD-In. I usually use the same text that will be displayed in the ADD-IN menu.

6) Lastly, notice that the ADD-IN version does not implement the XP-style button interface as the stand-alone version does. This is because the ADD-IN version is dependent upon the VB6 IDE for visual interface guidance. Because VB6 is actually designed to support the XP interface, you can copy the NewAPIViewer.exe.MANIFEST file to your Visual Studios\VB98 folder, rename it VB6.exe.MANIFEST, and you will enable XP-button display within VB6 (at a cost to some color selections, which I do not find a bother, and you might notice that it is the common controls version 5, not 6, that support some of the extended XP styles, such as progress bars, and you will find you will need to place radio buttons on pickturebox backgrounds if you wish them to display properly within frame controls).
