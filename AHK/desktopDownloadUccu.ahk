#SingleInstance force

;Variables
;The login page has one type of banner. All the pages after that have a maintenance banner.
;Login page banner is y93 to y237. 144 pixels tall. 
;Maintenance banner is y162 to y232. 70 pixels tall.

loginBannerColor := "0x219FB5" ;If login banner is there, it will be this color at this coordinates
xLoginBanner := 2844
yLoginBanner := 128
xSignInButton := 2200
ySignInButon := 474
littleBannerColor := "0x001E3C"
xLittleBanner := 1968
yLittleBanner := 103
xCloseLittleBanner := 1687 ;This is a bogus value. When the banner appears again, get the coordinates.
yCloseLittleBanner := 103
xRewards := 1865
yRewards := 707
xChecking := 1862
yChecking := 816
xExport := 2250
yExport := 368
xCsv := 2250
yCsv := 462
xBack := 1362
yBack := 154


Run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
Sleep 8000

color := PixelGetColor(xLoginBanner, yLoginBanner)
if(color = loginBannerColor){ ;If login page has banner
    Click xSignInButton, ySignInButon + 144
}else{ ;If no banner
    Click xSignInButton, ySignInButon
}
Sleep 9000


;If other page has banner, close it.
color := PixelGetColor(xLittleBanner, yLittleBanner)
if(color = littleBannerColor){
    Click xCloseLittleBanner, yCloseLittleBanner
}
Sleep 1000

;Rewards
Click xRewards, yRewards
Sleep 7000
Click xExport, yExport
Sleep 1000
Click xCsv, yCsv
Sleep 7000
Click xBack, yBack
Sleep 5000

;Checking
Click xChecking, yChecking
Sleep 7000
Click xExport, yExport
Sleep 1000
Click xCsv, yCsv
Sleep 7000

Send "^w"

Exit