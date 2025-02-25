# Panzer-Storage-Gauge-VB6
 
  A FOSS Storage Gauge VB6 WoW64 Widget for Reactos, XP, Win7, 8 and 10/11+.
 
 My current VB6/RC6 PSD program being finished now, fundamentally complete, only awaiting addition of smooth rotation to the Storage pointer, testing on Windows XP and Win7 32bit and some multi monitor checking, completion of the CHM help file and the creation of the setup.exe. This Panzer widget is based upon the Yahoo widget of the same visual design and very similar operation.

 Why VB6? Well, with a 64 bit, modern-language improvement upgrade on the way with 100% compatible TwinBasic coupled with support for transparent PNGs via RC/Cairo, VB6 code has an amazing future.

![vb6-logo-350](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/00e1e934-2d1e-447a-a40f-519aa49d5ef1)

 I created as a variation of the previous gauges I had previously created for the World of Tanks and War Thunder 
 communities. The Panzer Storage Gauge widget is an attractive dieselpunk VB6 widget for your desktop. 
 Functional and gorgeous at the same time. The graphics are my own, I took original inspiration from a clock face by Italo Fortana combining it with an aircraft gauge. It is all my code with some help from the chaps at VBForums (credits given). 
  
The Panzer Storage Gauge VB6  is a useful utility displaying the Storage usage of a selected drive from your system in a dieselpunk fashion on your desktop. This Widget is a moveable widget that you can move anywhere around the desktop as you require.

The following is the code used to extract the drive information from the running system via the GetDiskFreeSpace API. These are the pertinent bits:

    Private Declare Function GetDiskFreeSpace _
        Lib "kernel32" _
        Alias "GetDiskFreeSpaceA" ( _
            ByVal lpRootPathName As String, _
            lpSectorsPerCluster As Long, _
            lpBytesPerSector As Long, _
            lpNumberOfFreeClusters As Long, _
            lpTotalNumberOfClusters As Long) _
            As Long
                    
    Private storagePercent As Long
    Private storageFree As Double
    
    Private Function getDrivePercent(ByVal sRootDrive As String) As Long
        Dim lSecPerClust        As Long
        Dim lBytesPerSec        As Long
        Dim lFreeClusters       As Long
        Dim lTotalClusters      As Long
        
        Dim dTotalBytes         As Double
        Dim dFreeBytes          As Double
        Dim dUsedBytes          As Double
        Dim thePercentage       As Long
    
        Call GetDiskFreeSpace(sRootDrive, lSecPerClust, lBytesPerSec, lFreeClusters, lTotalClusters)
                            
        ' Get Total Disk Space, divided by 10000 to overcome the limitation of VB6 double var type for current and future large drives.
        ' after all, we only need the final percentage...
        
        dTotalBytes = CDbl(lBytesPerSec * CDbl(lSecPerClust * (lTotalClusters / 10000)))
        
        If dTotalBytes = 0 Then
            getDrivePercent = 0
            Exit Function
        End If
          
        ' Get Free Disk Space
        
        Dim a As Double
        a = CDbl(lSecPerClust) * CDbl(lFreeClusters)
        dFreeBytes = CDbl(lBytesPerSec * a / 10000)
                                
        ' Get Used Disk Space
        dUsedBytes = dTotalBytes - dFreeBytes
        
        ' get percentage of used space
        thePercentage = dUsedBytes / dTotalBytes * 100
        
        getDrivePercent = thePercentage  
    End Function    
    
    Private Function getDriveFree(ByVal sRootDrive As String) As Double
        Dim lSecPerClust        As Long
        Dim lBytesPerSec        As Long
        Dim lFreeClusters       As Long
        Dim lTotalClusters      As Long
        
        Dim dFreeGig            As Double
        Dim dUsedGig            As Double
    
        Call GetDiskFreeSpace(sRootDrive, lSecPerClust, lBytesPerSec, lFreeClusters, lTotalClusters)
                                                    
        ' Get Free Disk Space
        dFreeGig = CDbl(lBytesPerSec) * CDbl(lSecPerClust) * CDbl(lFreeClusters)
        dFreeGig = dFreeGig / 1024 ' KB
        dFreeGig = dFreeGig / 1024 ' MB
        dFreeGig = dFreeGig / 1024 ' GB
                                
        getDriveFree = dFreeGig
    End Function
    
 There are some other bits of code used for formatting the results that you will need. Just dig into the code of this desktop gauge to find the above and it will become apparent which other bits you require.
 Hope the code is useful to anyone else building system metric utilities using VB6/VBS/VBA/TwinBasic.
 
 ![panzer-storagephoto-1000](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/f7f27729-f2cc-4935-8382-34d8da8dde88)

 This widget can be increased in size, animation speed can be changed, 
 opacity/transparency may be set as to the users discretion. The widget can 
 also be made to hide for a pre-determined period.

![icon](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/8f2ece67-6c27-428a-a5ed-c3ff57d3cd9f)

 Right clicking will bring up a menu of options. Double-clicking on the widget will cause a personalised Windows application to 
 fire up. The first time you run it there will be no assigned function and so it 
 will state as such and then pop up the preferences so that you can enter the 
 command of your choice. The widget takes command line-style commands for 
 windows. Mouse hover over the widget and press CTRL+mousewheel up/down to resize. It works well on Windows XP 
 to Windows 11.
 
![tank-help-750](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/15e40483-341e-49c9-a1bb-7d8e934f218a)

 The Panzer Storage Gauge VB6 gauge is Beta-grade software, under development, not yet 
 ready to use on a production system - use at your own risk.

 This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS 
 project creating a WoW64 widget for the desktop. 
 
![Licence002](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/507eaed1-4d72-4e54-943b-3343d443eb52)

 It is open source to allow easy configuration, bug-fixing, enhancement and 
 community contribution towards free-and-useful VB6 utilities that can be created
 by anyone. The first step was the creation of this template program to form the 
 basis for the conversion of other desktop utilities or widgets. A future step 
 is new VB6 widgets with more functionality and then hopefully, conversion of 
 each to RADBasic/TwinBasic for future-proofing and 64bit-ness. 

![menu](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/14145e20-66bd-4441-9f7e-040832c9d876)

 This utility is one of a set of steampunk and dieselpunk widgets. That you can 
 find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery
 
 I do hope you enjoy using this utility and others. Your own software 
 enhancements and contributions will be gratefully received if you choose to 
 contribute.

![security](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/b9be6695-96ef-40fe-9803-e305d92b406a)


 BUILD: The program runs without any Microsoft plugins.
 
 Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
 Framework 2.2 & Rubberduck 2.4.1, RichClient 6
 
 Links:
 
	https://www.vbrichclient.com/#/en/About/
	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	VBAdvance  
 
 
 Tested on :
 
	ReactOS 0.4.14 32bit on virtualBox    
	Windows 7 Professional 32bit on Intel    
	Windows 7 Ultimate 64bit on Intel    
	Windows 7 Professional 64bit on Intel    
	Windows XP SP3 32bit on Intel    
	Windows 10 Home 64bit on Intel    
	Windows 10 Home 64bit on AMD    
	Windows 11 64bit on Intel  
 
![dock](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/19e3eebc-76d8-4dea-a69c-dc93ce0a4da4)

 CREDITS:
 
 I have really tried to maintain the credits as the project has progressed. If I 
 have made a mistake and left someone out then do forgive me. I will make amends 
 if anyone points out my mistake in leaving someone out.
 
 MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and 
 early 2000s. Thanks for VB6.
 
 Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for 
 VB6. Specifically the components using transparency and reading images directly 
 from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for 
 all VB6 native controls and giving us access to advanced features on controls 
 such as transparency.
 
 Shuja Ali @ codeguru for his settings.ini code.
 
 ALLAPI.COM        For the registry reading code.
 
 Rxbagain on codeguru for his Open File common dialog code without a dependent 
 OCX - http://forums.codeguru.com/member.php?92278-rxbagain
 
 si_the_geek       for his special folder code
 
 Elroy on VB forums for the balloon tooltips
 
 Harry Whitfield for his quality testing, brain stimulation and being an 
 unwitting source of inspiration. 
 
 Dependencies:
 
 o A windows-alike o/s such as Windows XP, 7-11 or Apple Mac OSX 11. 
 
 o Microsoft VB6 IDE installed with its runtime components. The program runs 
 without any additional Microsoft OCX components, just the basic controls that 
 ship with VB6.  

![vb6-logo-200](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/23f13e44-5d7d-4d94-af88-7bbe7de60a34)


 * Uses the latest version of the RC6 Cairo framework from Olaf Schmidt.
 
 During development the RC6 components need to be registered. These scripts are 
 used to register. Run each by double-clicking on them.
 
	RegisterRC6inPlace.vbs
	RegisterRC6WidgetsInPlace.vbs
 
 During runtime on the users system, the RC6 components are dynamically 
 referenced using modRC6regfree.bas which is compiled into the binary.	
 
 
 Requires a PzStorage Gauge folder in C:\Users\<user>\AppData\Roaming\ 
 eg: C:\Users\<user>\AppData\Roaming\PzStorage Gauge
 Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzStorage Gauge
 The above will be created automatically by the compiled program when run for the 
 first time.
 
 Uses just one OCX control extracted from Krools mega pack (slider). This is part 
 of Krools replacement for the whole of Microsoft Windows Common Controls found 
 in mscomctl.ocx. The slider control OCX file is shipped with this package.
 
 * CCRSlider.ocx
 
 This OCX will reside in the program folder. The program reference to this OCX is 
 contained within the supplied resource file Panzer Storage Gauge Gauge.RES. It is 
 compiled into the binary.
 
 * OLEGuids.tlb
 
 This is a type library that defines types, object interfaces, and more specific 
 API definitions needed for COM interop / marshalling. It is only used at design 
 time (IDE). This is a Krool-modified version of the original .tlb from the 
 vbaccelerator website. The .tlb is compiled into the executable.
 For the compiled .exe this is NOT a dependency, only during design time.
 
 From the command line, copy the tlb to a central location (system32 or wow64 
 folder) and register it.
 
 COPY OLEGUIDS.TLB %SystemRoot%\System32\
 REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB
 
 In the VB6 IDE - project - references - browse - select the OLEGuids.tlb



 * SETUP.EXE - The program is currently distributed using setup2go, a very useful 
 and comprehensive installer program that builds a .exe installer. Youll have to 
 find a copy of setup2go on the web as it is now abandonware. Contact me
 directly for a copy. The file "install PzStorage Gauge 0.1.0.s2g" is the configuration 
 file for setup2go. When you build it will report any errors in the build.
 
 * HELP.CHM - the program documentation is built using the NVU HTML editor and 
 compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and 
 the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
 the job admirably. The HTML files exist alongside the compiled CHM file in the 
 HELP folder.
 
  Project References:

	VisualBasic for Applications  
	VisualBasic Runtime Objects and Procedures  
	VisualBasic Objects and Procedures  
	OLE Automation  
	vbRichClient6  
 
 ![prefs-about](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/1ddd46d6-f7dc-4fe9-89f0-6c00738ad346)

 LICENCE AGREEMENTS:
 
 Copyright © 2023 Dean Beedell
 
 In addition to the GNU General Public Licence please be aware that you may use 
 any of my own imagery in your own creations but commercially only with my 
 permission. In all other non-commercial cases I require a credit to the 
 original artist using my name or one of my pseudonyms and a link to my site. 
 With regard to the commercial use of incorporated images, permission and a 
 licence would need to be obtained from the original owner and creator, ie. me.

![about](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/5dd1683f-f323-4b4b-a8a2-82973a6d6ddb)

![Panzer-CPU-Gauge-onDesktop](https://github.com/yereverluvinunclebert/Panzer-Storage-Gauge-VB6/assets/2788342/85c31cb9-30d8-45b4-b08d-5669edda740a)







 
