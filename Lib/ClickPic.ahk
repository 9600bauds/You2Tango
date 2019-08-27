;ClickPic, PicExists by 9600bauds

PicExists(filename) {
    CoordMode, Pixel, Mouse, Screen
    ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, %filename%
    if (ErrorLevel = 2){
        MsgBox PicExists - Could not conduct the search for %filename%.
        return false
    }
    if (ErrorLevel = 0){
        return true
    }
    return false
}

ClickPic(filename, offsetX:=0, offsetY:=0) {  
    CoordMode, Pixel, Mouse, Screen
    ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, %filename%
    if (ErrorLevel = 2){
        MsgBox ClickPic - Could not conduct the search for %filename%.
        return false
    }
    if (ErrorLevel = 1){
        MsgBox ClickPic - Icon %filename% could not be found on the screen.
        return false
    }
    
    FoundX += offsetX
    FoundY += offsetY
    Click, %FoundX%, %FoundY%
    return true
}

WaitNotPic(filename, interval:=100, tries:=25) {
    CoordMode, Pixel, Mouse, Screen
    Loop, 50
    {
        ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenWidth%, %A_ScreenHeight%, %filename%
        if (ErrorLevel == 2){
            MsgBox WaitNotPic - Could not conduct the search for %filename%.
            return false
        }
        if (ErrorLevel == 1){
            return 1
        }
        Sleep, 25
    }
    return 0
}