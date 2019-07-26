;ClickPic, PicExists by 9600bauds

PicExists(filename) {
    CoordMode, Pixel, Mouse, Screen
    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %filename%
    if (ErrorLevel = 2){
        MsgBox Could not conduct the search for %filename%.
        return false
    }
    if (ErrorLevel = 0){
        return true
    }
    return false
}

ClickPic(filename, offsetX:=0, offsetY:=0) {  
    CoordMode, Pixel, Mouse, Screen
    ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %filename%
    if (ErrorLevel = 2){
        MsgBox Could not conduct the search for %filename%.
        return false
    }
    if (ErrorLevel = 1){
        MsgBox Icon %filename% could not be found on the screen.
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
        ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %filename%
        if (ErrorLevel == 2){
            MsgBox Could not conduct the search for %filename%.
            return false
        }
        if (ErrorLevel == 1){
            return 1
        }
        Sleep, 25
    }
    return 0
}

IsNum( str ) { ;Fuck AHK.
	if str is number
		return true
	return false
}

IsAlwaysOnTop( Window ) {
    WinGet, Estilo, ExStyle, %Window%
    Return (Estilo & 0x8) ; 0x8 is WS_EX_TOPMOST.
}