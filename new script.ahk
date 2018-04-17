
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir C:\Users\yawik\Google Drive\AHK\my projects\chrome steps recorder  ; Ensures a consistent starting directory.
SetBatchLines -1
#SingleInstance,Force
#Include Chrome.ahk
Tab:=Chm_Create_Instance("C:\Users\yawik\AppData\Local\Google\Chrome\User Data\ChromeBot") ;Create instance using default Profile
Chm_Navigate(Tab,"https://yedion.yvc.ac.il/yedion/fireflyweb.aspx") ;Navigate to a URL
Tab.Evaluate("
(Join;
document.getElementById('R1C1').value='203942677';
document.getElementById('R1C2').value='gyd885';
document.querySelector('BUTTON.btn-u.btn-block').click();
)")
Tab.WaitForLoad()
;********************Create instance and use a path to profile***********************************
Chm_Create_Instance(Profile_Path=""){
	if !(Profile_Path){
		FileCreateDir, ChromeProfile ;Make sure folder exists
		ChromeInst := new Chrome("ChromeProfile") ;Create a new Chrome Instance
	}Else{
		try ChromeInst := new Chrome(Profile_Path) ;Create for profile lookup prfile by putting this in your url in Chrome chrome://version/
	}
	return Tab := ChromeInst.GetTab() ;Connect to Active tab
}

;********************Navigate to page***********************************
Chm_Navigate(Tab,URL){
	Tab.Call("Page.navigate", {"url": URL}) ;Navigate to URL
	Tab.WaitForLoad() ;Wait for page to finish loading
}
esc::
Tab.Call("Browser.close")
Tab.Disconnect()
ExitApp