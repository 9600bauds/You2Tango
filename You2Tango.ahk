#Include <ClickPic>

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force ; Close old versions of this script automatically.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

CopiarUnidadMedidaVentas(){
    if(not WinExist("ACTUALIZACION DE ARTICULOS")){
        MsgBox No existe ACTUALIZACION DE ARTICULOS.
        return
    }
    
    ;TEdit4 es el campo de unidad de Medida Ventas.
    ControlGetText, Clipboard, TEdit4, ACTUALIZACION DE ARTICULOS
}

PegarPrecio98o99(mult:=1){
    Clipboard := RegExReplace(Clipboard, ",", ".") ;Reemplazar comas por puntos.
    Clipboard := RegExReplace(Clipboard, "\.(?![^.]+$)")  ;Quitar todos los puntos excepto el último.
    Clipboard := RegExReplace(Clipboard, "[^0-9.]") ;Eliminar todo excepto números y puntos.
    if(not IsNum(Clipboard)) {
        MsgBox, Clipboard is not a number.
        return
    }
    multiplied := (Clipboard * mult)
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.
   
    if(not WinExist("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        MsgBox No existe ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO.
        return
    }
    
    If(!IsAlwaysOnTop("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        WinActivate, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
        WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    }
    
    WinMenuSelectItem, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO, , Modificar
    WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    if(PicExists("Images/ActualizacionPrecios/Dolar.png")){
        if(not ClickPic("Images/ActualizacionPrecios/Dolar.png", 425, 5)){
            return
        }
        WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
        if(not ClickPic("Images/ActualizacionPrecios/Dolar_Seleccionado.png", 425, 5)){
            return
        }
    }
    else{
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno.png", 425, 5)){
            return
        }
        WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno_Seleccionado.png", 425, 5)){
            return
        }
    }
    WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    
    ControlGetText, itemID, TEdit6, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    ControlGetText, oldPrice, TNumEditTg1, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    LogPriceChange(itemID, oldPrice, multiplied)
    
    Send, %multiplied%
    Send, {F10}
    Sleep, 150
    Send, {F10}
}

ActualizarDescripFecha(doAfter:="", replacement:=""){
    if(not WinExist("ACTUALIZACION DE ARTICULOS")){
        MsgBox No existe ACTUALIZACION DE ARTICULOS.
        return
    }
    
    if(replacement == ""){
        FormatTime, replacement, , MM/yyyy
    }
    
    if(WinExist("ahk_class TFrmBuscar")){
        CerrarVentanaBuscar()
    }
    
    If(!IsAlwaysOnTop("ACTUALIZACION DE ARTICULOS")){
        WinActivate, ACTUALIZACION DE ARTICULOS
        WinWait, ACTUALIZACION DE ARTICULOS
    }
    
    ;TEdit8 es la ID del campo de texto de Descripción Adicional.
    ControlGetText, Clipboard, TEdit8, ACTUALIZACION DE ARTICULOS ;Copiamos al portapapeles, por si accidentalmente sobreescribimos la descripción de un artículo equivocado.
    WinMenuSelectItem, ACTUALIZACION DE ARTICULOS, , Modificar
    WinWait, ACTUALIZACION DE ARTICULOS
    ControlFocus, TEdit8, ACTUALIZACION DE ARTICULOS ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    ControlSetText, TEdit8, %replacement%, ACTUALIZACION DE ARTICULOS
    WinWait, ACTUALIZACION DE ARTICULOS
    Send, {F10}
    Sleep, 150
    Send, {F10}
    Sleep, 150
    Send, {F10}
    
    if(doAfter == "search"){
        Sleep, 150
        WinMenuSelectItem, ACTUALIZACION DE ARTICULOS, , Buscar, Por Clave
    }
    else if(doAfter == "next"){
        Sleep, 150
        ProximoArticulo()
    }
}

BuscarPorPortapapel(){
    if WinExist("OpenOffice Calc"){
        WinActivate, OpenOffice Calc
        WinWait, OpenOffice Calc
        if WinExist("Find & Replace"){
            WinActivate, Find & Replace
            WinWait, Find & Replace
            Send, !s ;Alt+S: Search For
        }
        else{
            Send, ^f ;Ctrl+F: Buscar
        }
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
        Sleep, 100
        if(PicExists("Images/OpenOfficeCalc/EndOf.png")){ ;Damn you, OpenOffice.
            Send, {Enter}
        }
    }
    if WinExist("Adobe Reader"){
        WinActivate, Adobe Reader
        WinWait, Adobe Reader
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
    }
}

SincronizarArticulosPrecio(){
    if(not WinExist("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        MsgBox No existe ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO.
        return
    }
    if(not WinExist("ACTUALIZACION DE ARTICULOS")){
        MsgBox No existe ACTUALIZACION DE ARTICULOS.
        return
    }
    if(WinExist("ahk_class TFrmBuscar")){
        CerrarVentanaBuscar()
    }

    ControlGetText, CodigoArticulo, TEdit11, ACTUALIZACION DE ARTICULOS ;TEdit11 es el campo de código de artículo.
    WinMenuSelectItem, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO, , Buscar, Por Clave
    WinWait, ahk_class TFrmBuscar ;Ésta es la ventana Buscar.
    ControlSend, TcxCustomInnerTextEdit1, %CodigoArticulo%, ahk_class TFrmBuscar 
    
    CerrarVentanaBuscar()
}

CerrarVentanaBuscar(){
    if(not WinExist("ahk_class TFrmBuscar")){
        MsgBox No existe ahk_class TFrmBuscar.
        return
    }
    WinActivate, ahk_class TFrmBuscar
    
    Control, Check, , TCheckBox4, ahk_class TFrmBuscar ;Activar FILTRAR
    Control, Uncheck, , TCheckBox3, ahk_class TFrmBuscar ;Desactivar INCREMENTAL
    
    ControlGetText, CodigoIngresado, TcxCustomInnerTextEdit1, ahk_class TFrmBuscar
    if(CodigoIngresado == ""){
        Send, {Esc}
        WinWaitClose, ahk_class TFrmBuscar
        return
    }
    
    Send, {Enter}
    If(WinExist("ahk_class TFrmBuscar")){ ;Puede que ya hayamos apretado Enter nosotros.
        Send, {Enter}
    }
    WinWaitClose, ahk_class TFrmBuscar
}

ProximoArticulo(){ 
    if(WinExist("ACTUALIZACION DE ARTICULOS")){
        ;WinWait, ACTUALIZACION DE ARTICULOS
        WinMenuSelectItem, ACTUALIZACION DE ARTICULOS, , Buscar, Siguiente ;Modo Modificar.
        ;ControlSend,,{PGDN}, ACTUALIZACION DE ARTICULOS
    }
    if(WinExist("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        ;WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
        WinMenuSelectItem, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO, , Buscar, Siguiente
        ;ControlSend,,{PGDN}, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    }
}

AnteriorArticulo(){
    if(WinExist("ACTUALIZACION DE ARTICULOS")){
        ControlSend,,{PGUP}, ACTUALIZACION DE ARTICULOS
    }
    if(WinExist("ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO")){
        ControlSend,,{PGUP}, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    }
}

EstilizarVentanas(Activar := 1){
    if(Activar == 1){
        WinSet, AlwaysOnTop, On, ACTUALIZACION DE ARTICULOS
        WinSet, Region, 0-0 W572 H222, ACTUALIZACION DE ARTICULOS ;Máscara de 572x222 empezando en 0,0
        WinMove, ACTUALIZACION DE ARTICULOS, , 1028, 26, 572
        WinGetPos, X, Y, W, H, ACTUALIZACION DE ARTICULOS
        
        WinSet, AlwaysOnTop, On, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO

        WinMove, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO, , X, Y+222, W, H
        WinGetPos, X, Y, W, H, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
        
        if(WinExist("ahk_class Notepad")){
            WinSet, Region, 0-0 W572 H398, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO ;Máscara de 572x425 empezando en 0,0
            
            WinSet, Region, 0-0 W999 H999, ahk_class Notepad ;Literalmente sólo para que tenga los 3 pixeles negros feos
            WinSet, AlwaysOnTop, On, ahk_class Notepad
            WinMove, ahk_class Notepad, , X, Y+398 , W, 125
        }
        

    }
    else{
        WinSet, AlwaysOnTop, Off, ACTUALIZACION DE ARTICULOS
        WinSet, Region, , ACTUALIZACION DE ARTICULOS
        
        WinSet, AlwaysOnTop, Off, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
        WinSet, Region, , ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
        
        WinSet, AlwaysOnTop, Off, ahk_class Notepad
        WinSet, Region, , ahk_class Notepad
    }
}

LogPriceChange(itemID := "", oldPrice := "", newPrice = ""){
    percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    finalText = %itemID%: %oldPrice% -> %newPrice% (%percent%`%)`r`n
    Control, EditPaste, %finalText%, , ahk_class Notepad
}

Launch_Media::
MsgBox, Testing...
return

Volume_Up::
EstilizarVentanas(1)
return

Volume_Down::
EstilizarVentanas(0)
return

;Volume_Mute::

Media_Play_Pause::
ActualizarDescripFecha("search")
return

^Media_Play_Pause::
ActualizarDescripFecha("next")
return

Media_Prev::
AnteriorArticulo()
return

Media_Next::
ProximoArticulo()
return

Launch_Mail::
SincronizarArticulosPrecio()
return

^Launch_Mail::
SincronizarArticulosPrecio()
Sleep, 100
CopiarUnidadMedidaVentas()
Sleep, 100
BuscarPorPortapapel()
return

Browser_Search::
CopiarUnidadMedidaVentas()
Sleep,100
BuscarPorPortapapel()
return

Browser_Home::
PegarPrecio98o99(1.21)
return

^Browser_Home::
PegarPrecio98o99(1.21/45)
return

#IfWinActive SOS DE STOCK ; Works for EGRESOS and INGRESOS. AHK does not have an OR operand for this command.
::cdm::Cambio de Mercadería - Blas
::cds::Corrección de Stock - Blas
::mui::Uso Interno - Blas
::-b:: - Blas
#IfWinActive