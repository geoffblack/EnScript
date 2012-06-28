Every once in a while I see people looking for help building dll's in the Microsoft .Net Framework for use with EnScript through COM. There are some tricks that are definitely non-obvious, so I put together this package to help. Hopefully you find it useful.

Any questions or problems, email me: 

Geoff Black
geoff@geoffblack.com
http://www.geoffblack.com

==============

There are a few problems people commonly run into: 

1: When you're generating the TypeLib using regasm.exe it's generating a default interface with no methods because of the way .NET works. The default generated interface will always be the class name with an underscore in front of it (_Main). If you examine the generated tlb file using the OLE/COM Object Viewer from Visual Studio, Tools -> OLE/COM Object Viewer then File -> View TypeLib, I think you'll find it shows no methods under the default interface. You need to explicitly define the interface. VBScript makes only late-bound calls into the assembly; it doesn't use the typelib so it doesn't care that the default interface is empty. EnCase is making an early-bound call (as would, e.g. VB6) using the defined typelib which has no methods the way you're doing it.

2. RegAsm.exe doesn't add all of the keys in the registry that it should be adding for early-binders (this is by design).

Of course it's always best to (1) use strong-named assemblies (signed), (2) generate a unique GUID defined in the code so regasm isn't generating a new one every time you run it (run guidgen.exe to generate your own - one for the interface and another one for the class), and (3) install the assembly into the GAC (Global Assembly Cache) with gacutil before using regasm.

==============

This package contains a demo EnScript and the Visual C#.NET project has recently been updated to VS2010 format with a target build of .Net Framework 4. I created a reg file in the CDemoLib\bin\Release folder as well as batch files to register and unregister the assembly. You may have to change the location of RegAsm.exe and gacutil.exe to suit your local box since I hard coded the paths into the batch file.

Run "reg_Net.v*.x_**bitOS.bat.bat" (depending on your target Framework and OS) then "COM_obj.reg".

***The reg file will only work with the attached assembly. As you will be generating your own GUIDs, you will have to create the keys under your specific classID GUID.

==============

What the registry file does (how to do it manually) and why you need it:

As I stated earlier, RegAsm.exe doesn't add all of the keys in the registry that it should be adding for early-binders (this is by design).

Open regedit and go to HKCR\CLSID\[your class GUID]. You'll need to add a Typelib key here and then set the Default value to the GUID of your Typelib - {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}. Also add a Control key with no values, a MiscStatus key with a Default value of 131457, and a Version key with the Typelib version (probably 1.0 unless you changed it or added more than one but check your Typelib).

After you do this, you will be able to compile the typelib into an EnScript and afterwards when you View -> EnScript Types you'll see your imported type listed there.

This is when it comes in very handy to have pregenerated your GUIDs (one for the interface, one for the class, and one for the assembly). The GUID used in the CLSID tree is the one you assigned to the class. The GUID used in the Typelib tree is the one you assigned to the assembly.

==============

FAQ:

Q: Where did you generate the "COM_Obj.reg" file from? Was it a manual extraction from the registry (How did it get there?) or did I only need to do that because it was compiled on another machine and the compiler on that machine originally made the registry entries?

A: I manually created the registry entries. RegAsm will not make them for you. Checking "Register for COM Interop" or "Make visible to COM" in your VS project simply automates RegAsm with the /codebase switch. If you used Windows Installer to distribute the component, you could obviously specify reg keys to add in the installer. I don't know of any other way to do it, but I'm sure there has to be a better way.

Regarding COM_obj.reg, it is specific to the GUID for that class and assembly so you would not be able to reuse it without changing the two GUIDs. 

-------------

Q: The 2 GUIDs that are in the reg file are the GUID that you assigned to the class and the other GUID is the one that C#.Net assigns to the assembly (click on the projects property page and then the assembly tab). Is this correct?

A: The assembly GUID can be added manually into AssemblyInfo.cs or from the Project info page, click the big button that says Assembly Info or something like that and there's a field for GUID. The class GUID is assigned in the code directly preceding the beginning of the class.

You should always generate your own unique GUIDs using guidgen.exe and place them in the project before compiling and registering, this way you'll always have the same GUID and you'll know where to look in the registry. Otherwise, Visual Studio will generate a new one for the assembly and class every time you compile and register.

-------------

(Un)Common EnCase COM Knowledge

EnCase doesn't refresh the typelib every time you compile an EnScript. This only occurs the first time. If you change anything in the library, you should close EnCase and restart it, then compile your EnScript again. You'll be able to see it in the EnScript Types tab.
