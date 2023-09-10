#SingleInstance force

;Variables
;The login page has one type of banner. All the pages after that have a maintenance banner.
;Login page banner is y93 to y237. 144 pixels tall. 
;Maintenance banner is y162 to y232. 70 pixels tall.

loginBannerColor := "0x08A3B3" ;If login banner is there, it will be this color at this coordinates
xLoginBanner := 1762
yLoginBanner := 191
xSignInButton := 1458
ySignInButon := 451
littleBannerColor := "0x001E3C"
xLittleBanner := 1599
yLittleBanner := 213
xCloseLittleBanner := 1687
yCloseLittleBanner := 201
xRewards := 1134
yRewards := 896
xChecking := 1134
yChecking := 1019
xExport := 1609
yExport := 496
xCsv := 1609
yCsv := 612
xBack := 507
yBack := 194


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