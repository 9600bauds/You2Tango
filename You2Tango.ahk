#Include <ClickPic>

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

CopiarUnidadMedidaVentas(){
    if(not WinExist("ACTUALIZACION DE ARTICULOS")){
        MsgBox No existe ACTUALIZACION DE ARTICULOS.
        return
    }
    
    WinMenuSelectItem, ACTUALIZACION DE ARTICULOS, , Modificar ;Modo Modificar.
    WinWait, ACTUALIZACION DE ARTICULOS
    
    ControlFocus, TEdit4, ACTUALIZACION DE ARTICULOS ;TEdit4 es el ID del campo de texto Unidad de Medida Ventas.    
    SendMessage, 0x301, , , TEdit4, ACTUALIZACION DE ARTICULOS ;"SendMessage, 0x301" envia CTRL+C.
    WinWait, ACTUALIZACION DE ARTICULOS
    
    ControlSend,,{Esc}, ACTUALIZACION DE ARTICULOS ;Cancelar modo Modificar
    WinWait, ACTUALIZACION DE ARTICULOS
}

PegarPrecio98o99(mult:=1){
    Clipboard := RegExReplace(Clipboard, "\R") ;Eliminar líneas extra generadas por Excel y otros programas
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
    WinActivate, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    
    WinWait, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO
    WinMenuSelectItem, ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO, , Modificar
    WinWait, ACTUALIZACION DE ARTICULOS
    if(PicExists("Images/ActualizacionPrecios/Dolar.png")){
        if(not ClickPic("Images/ActualizacionPrecios/Dolar.png", 425, 5)){
            return
        }
        WinWait, ACTUALIZACION DE ARTICULOS
        if(not ClickPic("Images/ActualizacionPrecios/Dolar_Seleccionado.png", 425, 5)){
            return
        }
    }
    else{
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno.png", 425, 5)){
            return
        }
        WinWait, ACTUALIZACION DE ARTICULOS
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno_Seleccionado.png", 425, 5)){
            return
        }
    }
    WinWait, ACTUALIZACION DE ARTICULOS
    Send, %multiplied%
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
    
    WinActivate, ACTUALIZACION DE ARTICULOS
    WinWait, ACTUALIZACION DE ARTICULOS
    
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
    }
    if WinExist("Adobe Reader"){
        WinActivate, Adobe Reader
        WinWait, Adobe Reader
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
    }
}

;Return TRUE: Se completó con éxito, incluyendo apretar enter en la búsqueda
;Return FALSE: Probablemente no se completó la búsqueda y es necesario apretar Enter.
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
    
    return CerrarVentanaBuscar()
}

CerrarVentanaBuscar(){
    if(not WinExist("ahk_class TFrmBuscar")){
        MsgBox No existe ahk_class TFrmBuscar.
        return
    }
    WinActivate, ahk_class TFrmBuscar
    
    ControlGetText, CodigoIngresado, TcxCustomInnerTextEdit1, ahk_class TFrmBuscar
    if(CodigoIngresado == ""){
        Send, {Esc}
        WinWaitClose, ahk_class TFrmBuscar
        return true
    }
    
    ControlGetText, ModoDeBusqueda, TComboBox1, ahk_class TFrmBuscar
    if(ModoDeBusqueda == "Sinónimo"){
        ImageToSearch = Images/VentanaBuscar/Sinonimo_Inactivo.png
        ControlGetFocus, CampoSeleccionado, ahk_class TFrmBuscar
        if(CampoSeleccionado == "TcxCustomInnerTextEdit1"){
            ImageToSearch = Images/VentanaBuscar/Sinonimo_Activo.png
        }
    }
    else if(ModoDeBusqueda == "Cód. Artículo"){
        ImageToSearch = Images/VentanaBuscar/CodArticulo_Inactivo.png
        ControlGetFocus, CampoSeleccionado, ahk_class TFrmBuscar
        if(CampoSeleccionado == "TcxCustomInnerTextEdit1"){
            ImageToSearch = Images/VentanaBuscar/CodArticulo_Activo.png
        }
    }
        
    if(WaitNotPic(ImageToSearch)){
        Send, {Enter}
        WinWaitClose, ahk_class TFrmBuscar
        return true
    }
    Send, {Esc}
    WinWaitClose, ahk_class TFrmBuscar
    return false
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

Media_Prev::
ActualizarDescripFecha("search")
return

^Media_Prev::
ActualizarDescripFecha("next")
return

Media_Next::
ProximoArticulo()
return

^Media_Next::
AnteriorArticulo()
return

Launch_Mail::
SincronizarArticulosPrecio()
return

^Launch_Mail::
if(SincronizarArticulosPrecio() == true){
    Sleep, 800
    CopiarUnidadMedidaVentas()
    Sleep,100
    BuscarPorPortapapel()
}
return

Browser_Search::
CopiarUnidadMedidaVentas()
return

^Browser_Search::
CopiarUnidadMedidaVentas()
Sleep,100
BuscarPorPortapapel()
return

Browser_Home::
WinActivate, ACTUALIZACION DE ARTICULOS ;HITLERS
PegarPrecio98o99(1.16)
return

^Browser_Home::
WinActivate, ACTUALIZACION DE ARTICULOS ;HITLERS
PegarPrecio98o99(1.04975)
return

#IfWinActive SOS DE STOCK ; Works for EGRESOS and INGRESOS. AHK does not have an OR operand for this command.
::cdm::Cambio de Mercadería - Blas
::cds::Corrección de Stock - Blas
::mui::Uso Interno - Blas
::-b:: - Blas
#IfWinActive