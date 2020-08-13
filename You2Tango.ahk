#Include <ClickPic>

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance force ; Close old versions of this script automatically.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2 ; Match window titles anywhere, not just at the start.

;{ Globals (most of these are effectively defines)
global ventanaArticulos := "ACTUALIZACION DE ARTICULOS"
global campoMedidaVentas := "TEdit4"
global campoCodigoArt_Articulos := "TEdit11"
global campoDescAdicional := "TEdit8"
global campoDesc_Articulos := "TEdit9"
global campoCodBarra := "TEdit6"

global ventanaArticulos_OKchild := "ahk_class TFormChildTg"
global campoOKChild_1 := "TNumEditTg4" ;the long awaited comeback to OK boomer
global campoOKChild_2 := "TEdit4"

global ventanaProveedor := "Proveedores Asociados"
global campoCodigoProveedor := "TEdit3"

global ventanaPrecios := "ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO"
global campoCodigoArt_Precios_ModoNoModificar := "TEdit2" ;EN MODO NO MODIFICAR
global campoCodigoArt_Precios_ModoModificar := "TEdit6" ;EN MODO MODIFICAR
global campoPrecioActual := "TNumEditTg1"

global ventanaBuscar := "ahk_class TFrmBuscar"
global ventanaBuscar_Articulos := "ahk_class TFrmBuscar ahk_exe ST_1.EXE"
global ventanaBuscar_Precios := "ahk_class TFrmBuscar ahk_exe GV_1.EXE"
global campoContenido_Buscar := "TcxCustomInnerTextEdit1"
global checkboxFiltrar := "TCheckBox4"
global checkboxIncremental := "TCheckBox3"
global buscarEn := "TComboBox1"
global buscarCodigo := 1
global buscarSinonimo := 4

global ventanaIngresosStock := "INGRESOS DE STOCK"
global ventanaEgresosStock := "EGRESOS DE STOCK"

global ventanaNotepad := "ahk_class Notepad"
global ventanaCalc := "OpenOffice Calc"
global ventanaCalc_Buscar := "Find & Replace"
global ventanaCalc_Main := "ahk_class SALFRAME" ; Precisamente la planilla principal, no ningún diálogo
global ventanaAdobeReader := "Adobe Reader"
global ventanaAdobeReader_Buscar := "ahk_class AVL_AVWindow"
global ventanaAdobeReader_BuscarOK := "Button18"
global ventanaAdobeReader_Buscar_Input := "Edit4"
global ventanaAbodeReader_Buscar_Matches := "Static12"

global multiplicadorPrecio1 := 1.21
global multiplicadorPrecio2 := 1
global multiplicadorExtra := 0

global search_Default = "Default"
global search_Exact = "Exact"
global search_Start = "Match Start"
global search_End = "Match End"
global search_RemoveLastWord = "Remove Last Word"
global search_LongestNumber = "Longest Number"
global search_Fabrimport = "Fabrimport"
global search_Faroluz = "Faroluz"
global search_Ferrolux = "Ferrolux"
global search_Solnic = "Solnic"
global searchType := "Default"

global parseNoDecimals := false

global enableCodeArray := true
global codeArray := []

global AdHocMode := false

global PostSearchString := ""
;}

;{ Ventana Artículos - Helpers
GetUnidadMedidaVentas(noModify := false){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }
       
    WinWait, %ventanaArticulos%
    ControlGetText, unidadMedida, %campoMedidaVentas%, %ventanaArticulos%
    
    if(unidadMedida == "" or unidadMedida == "NO TRAER" or RegExMatch(unidadMedida, "$[-]+^")){
        MsgBox, GetUnidadMedidaVentas - Unidad inválida. (%unidadMedida%)
        return
    }
    
    if(!noModify){
        if(searchType == search_Exact){
            unidadMedida := "^" . unidadMedida . "$"
        }
        else if(searchType == search_Start){
            unidadMedida := "^" . unidadMedida
        }
        else if(searchType == search_End){
            unidadMedida := unidadMedida . "$"
        }
        else if(searchType == search_RemoveLastWord){
            unidadMedida := RegExReplace(unidadMedida, " \w+$", "")
        }
        else if(searchType == search_LongestNumber){
            longestMatch := ""
            for index, match in AllRegexMatches(unidadMedida, "[\d]+"){
                if(StrLen(match) > StrLen(longestMatch)){
                    longestMatch := match
                }
            }
            unidadMedida := longestMatch
        }
        else if(searchType == search_Fabrimport){
            if(InStr(unidadMedida, "*")){
                MsgBox, GetUnidadMedidaVentas - Salteando por asterisco. (%unidadMedida%)
                return
            }
            unidadMedida := "[^0-9]" . unidadMedida . "$"
        }
        else if(searchType == search_Faroluz){
            unidadMedida := RegExReplace(unidadMedida, " \w+$", "") . "$"
        }
        else if(searchType == search_Ferrolux){
            RegExMatch(unidadMedida, "([A-Z]+-\d+)", unidadMedida)
        }
        else if(searchType == search_Solnic){
            unidadMedida := "^" . unidadMedida . "[\s+|$]"
        }
    }
        
    return unidadMedida
}

CambiarCampoVentanaArticulos(field := "", newText = ""){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return false
    }
    
    ;if(WinExist(ventanaBuscar)){
    ;    CerrarVentanaBuscar()
    ;-}
    
    If(!IsAlwaysOnTop(ventanaArticulos)){
        WinActivate, %ventanaArticulos%
        WinWait, %ventanaArticulos%
    }
    
    WinMenuSelectItem, %ventanaArticulos%, , Modificar
    WinWait, %ventanaArticulos%
    ControlFocus, %field%, %ventanaArticulos% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    ControlSetText, %field%, %newText%, %ventanaArticulos%
    WinWait, %ventanaArticulos%
    AceptarCambiosVentanaArticulos()
    
    return true
}

GetCodigoVentanaArticulos(){
    if(not WinExist(ventanaArticulos)){
        MsgBox GetCodigoVentanaArticulos - No existe %ventanaArticulos%.
        return
    }
    
    ControlGetText, itemID, %campoCodigoArt_Articulos%, %ventanaArticulos%
    return itemID
}

AceptarCambiosVentanaArticulos(){
    DetectHiddenWindows, On
    
    if(not WinExist(ventanaArticulos)){
        MsgBox AceptarCambiosVentanaArticulos - No existe %ventanaArticulos%.
        return
    }
    
    ControlSend, %campoMedidaVentas%, {F10}, %ventanaArticulos%
    WinWait, %ventanaArticulos_OKchild%
    ControlSend, %campoOKchild_1%, {F10}, %ventanaArticulos_OKchild%
    WinWait, %ventanaArticulos_OKchild%
    ControlSend, %campoOKchild_2%, {F10}, %ventanaArticulos_OKchild%
GetIVAType(){
    DetectHiddenWindows, On
    
    if(not WinExist(ventanaArticulos)){
        MsgBox GetIVAType - No existe %ventanaArticulos%.
        return
    }
    
    WinMenuSelectItem, %ventanaArticulos%, , Modificar
    WinWait, %ventanaArticulos%
    ControlSend, %campoDesc_Articulos%, {F10}, %ventanaArticulos%
    WinWait, %ventanaArticulos_OKchild%
    ControlGetText, IVAType, %campoOKChild_1%, %ventanaArticulos_OKchild%
    ControlSend, %campoOKchild_1%, {Esc}, %ventanaArticulos_OKchild%
    WinWait, %ventanaArticulos%
    ControlSend, %campoDesc_Articulos%, {Esc}, %ventanaArticulos%
    WinWait, %ventanaArticulos%
    return IVAType
}

FixIVA(){
    ControlGetText, descAdicional, %campoDescAdicional%, %ventanaArticulos% 
    if(not InStr(descAdicional, "½IVA") and GetIVAType() == 2){
        ;MsgBox, Arreglando descripción adicional para incluir ½IVA.
        descfinal = %descAdicional% ½IVA
        CambiarCampoVentanaArticulos(descAdicional, descfinal)
        return
    }
}

GoToVentanaArticulos(num := "", searchType := 0){
    WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
    WinWait, %ventanaBuscar%
    if(searchType != 0){
        setSearchTypeVentanaBuscar(searchType)
    }
    
    ControlFocus, %campoContenido_Buscar%, %ventanaBuscar% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    Control, EditPaste, %num%, %campoContenido_Buscar%, %ventanaBuscar%

    CerrarVentanaBuscar()
}
;}

;{ Ventana Artículos - Funciones
EliminacionArticulo(doAfter:=""){    
    itemID := GetCodigoVentanaArticulos() ;Para el logging
    ControlGetText, oldDesc, %campoDesc_Articulos%, %ventanaArticulos% ;Para el logging.
    ControlGetText, Clipboard, %campoDesc_Articulos%, %ventanaArticulos% ;Copiamos al portapapeles, por si accidentalmente borramos un artículo equivocado.
    
    if(not CambiarCampoVentanaArticulos(campoDesc_Articulos, "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ")){
        return false
    }
    
    LogArticleDeletion(itemID, oldDesc)
       
    if(doAfter == "search"){
        Sleep, 150
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
    }
    else if(doAfter == "next"){
        Sleep, 150
        ProximoArticulo()
    }
}

ActualizarDescripFecha(doAfter:="", replacement:=""){   
    if(replacement == ""){
        FormatTime, replacement, , MM/yyyy
    }
        
    itemID := GetCodigoVentanaArticulos() ;Para el logging.
    ControlGetText, oldDesc, %campoDescAdicional%, %ventanaArticulos% ;Para el logging.
    ControlGetText, Clipboard, %campoDescAdicional%, %ventanaArticulos% ;Copiamos al portapapeles, por si accidentalmente sobreescribimos la descripción de un artículo equivocado.
    
    RegExMatch(oldDesc, "( .*)", extraInfo) ;Preserve everything after a space.
    replacement := replacement . extraInfo1
    
    if(not CambiarCampoVentanaArticulos(campoDescAdicional, replacement)){
        return false
    }
    
    LogDescChange(itemID, oldDesc, replacement)
    
    if(doAfter == "search"){
        Sleep, 150
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
    }
    else if(doAfter == "next"){
        Sleep, 150
        ProximoArticulo()
    }
}

MassActualizarDesc(){
    if(not WinExist(ventanaArticulos)){
        MsgBox MassActualizarDesc - No existe %ventanaArticulos%.
        return
    }
    
    explanation := "Ingrese la lista de códigos (completos, no sinónimos) a actualizar, delimitados por coma:"
    InputBox, arr, Actualizar Fechas, %explanation%
    StringReplace , arr, arr, %A_Space%,,All
        
    Loop, parse, arr, `,,
    {
        GoToVentanaArticulos(A_LoopField, buscarCodigo)
        Sleep, 250
        ActualizarDescripFecha()
        Sleep, 250
    }
}

CorregirUnidadMedidaVentas(prov := "", useClipboard := true){ ;Desvergonzadamente ad-hoc.
    if(not WinExist(ventanaArticulos)){
        MsgBox CorregirMedidaVentas - No existe %ventanaArticulos%.
        return false
    }
    
    If(!IsAlwaysOnTop(ventanaArticulos)){
        WinActivate, %ventanaArticulos%
        WinWait, %ventanaArticulos%
    }
    
    initialMedidaVentas := GetUnidadMedidaVentas(true)
    if(!useClipboard){
        Clipboard := initialMedidaVentas
    }
    if(prov == "Ferrolux"){
        Clipboard := RegExReplace(Clipboard, "([a-zA-Z])([1-9])","$1-$2")
        Clipboard := RegExReplace(Clipboard, " ","")
    }
    if(Clipboard != initialMedidaVentas){
        CambiarCampoVentanaArticulos(campoMedidaVentas, Clipboard)
    }
}

CambiarProveedor(prov := ""){
    if(prov == ""){
        return false
    }
    
    if(not WinExist(ventanaArticulos)){
        MsgBox CorregirMedidaVentas - No existe %ventanaArticulos%.
        return false
    }
    
    If(!IsAlwaysOnTop(ventanaArticulos)){
        WinActivate, %ventanaArticulos%
        WinWait, %ventanaArticulos%
    }
    
    WinMenuSelectItem, %ventanaArticulos%, , Proveedor
    WinWait, %ventanaProveedor%
    
    WinGetPos,,, proov_w, proov_h, %ventanaProveedor%
    WinMove, %ventanaProveedor%,, (A_ScreenWidth/2)-(proov_w/2), (A_ScreenHeight/2)-(proov_h/2)
    WinSet, AlwaysOnTop, On, %ventanaProveedor%
    
    ControlFocus, %campoCodigoProveedor%, %ventanaProveedor% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    ControlSetText, %campoCodigoProveedor%, %prov%, %ventanaProveedor%
    WinWait, %ventanaProveedor%
    ControlSend, %campoCodigoProveedor%, {Enter}, %ventanaProveedor%
    WinWait, %ventanaProveedor%
    ControlSend, %campoCodigoProveedor%, {F10}, %ventanaProveedor%
    WinWait, %ventanaProveedor%
    ControlSend, %campoCodigoProveedor%, {F10}, %ventanaProveedor%
}

PegarUnidadMedidaVentas(){ ;Para medidas desesperadas.
    if(not WinExist(ventanaArticulos)){
        MsgBox PegarUnidadMedidaVentas - No existe %ventanaArticulos%.
        return false
    }
    
    If(!IsAlwaysOnTop(ventanaArticulos)){
        WinActivate, %ventanaArticulos%
        WinWait, %ventanaArticulos%
    }
    
    ;todo logging
    Clipboard := RegExReplace(Clipboard, " ","")
    CambiarCampoVentanaArticulos(campoMedidaVentas, Clipboard)
}
;}

;{ Ventana Precios - Helpers
SeleccionarPrecio98o99(){
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return false
    }
    
    If(!IsAlwaysOnTop(ventanaPrecios)){
        WinActivate, %ventanaPrecios%
        WinWait, %ventanaPrecios%
    }
    
    WinMenuSelectItem, %ventanaPrecios%, , Modificar
    WinWait, %ventanaPrecios%
    if(PicExists("Images/ActualizacionPrecios/Dolar.png")){
        if(not ClickPic("Images/ActualizacionPrecios/Dolar.png", 425, 5)){
            return false
        }
        WinWait, %ventanaPrecios%
        if(not ClickPic("Images/ActualizacionPrecios/Dolar_Seleccionado.png", 425, 5)){
            return false
        }
    }
    else{
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno.png", 425, 5)){
            return false
        }
        WinWait, %ventanaPrecios%
        if(not ClickPic("Images/ActualizacionPrecios/NoUsarUsoInterno_Seleccionado.png", 425, 5)){
            return false
        }
    }
    WinWait, %ventanaPrecios%
    
    return true
}

GetCodigoVentanaPrecios(){
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return
    }
    
    ControlGetText, itemID, %campoCodigoArt_Precios_ModoNoModificar%, %ventanaPrecios%
    if(itemID == "S" or itemID == "N"){
        ControlGetText, itemID, %campoCodigoArt_Precios_ModoModificar%, %ventanaPrecios%
    }
    return itemID
}

GoToVentanaPrecios(num := ""){
    WinMenuSelectItem, %ventanaPrecios%, , Buscar, Por Clave
    WinWait, %ventanaBuscar%
    
    ControlFocus, %campoContenido_Buscar%, %ventanaBuscar% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    Control, EditPaste, %num%, %campoContenido_Buscar%, %ventanaBuscar%

    CerrarVentanaBuscar()
}
;}

;{ Ventana Precios - Funciones
IngresarMultiplicadoresPrecio(){
    explanation := "Ingrese el descuento con el siguiente formato:`n1.21 o 21% para +21`%`n0.8 o -20% para -20`%"
    newMultiplier := 0
    InputBox, newMultiplier, Descuento Básico, %explanation%,,,,,,,,%multiplicadorPrecio1%
    if ErrorLevel
        return ;Cancel
    if(ParsePercent(newMultiplier)) {
        multiplicadorPrecio1 := newMultiplier
    }
    else{
        MsgBox, No se ingresó un número. (%newMultiplier%)
    }
    
    InputBox, newMultiplier, Descuento Alternativo, %explanation%,,,,,,,,%multiplicadorPrecio2%
    if ErrorLevel
        return ;Cancel
    if(ParsePercent(newMultiplier)) {
        multiplicadorPrecio2 := newMultiplier
    }
    else{
        MsgBox, No se ingresó un número. (%newMultiplier%)
    }
}

PegarPrecio98o99(mult:=1){
    if(not SincronizadosArticulosPrecio()){
        MsgBox, PegarPrecio98o99 - Ventana Artículos y ventana Precios no están actualizadas!
        return false
    }
    
    mult := ParsePercent(mult)
    if(mult == 0){
        MsgBox, Multiplicador inválido! - PegarPrecio98o99
        return false
    }
    ControlGetText, descAdicional, %campoDescAdicional%, %ventanaArticulos% 
    RegExMatch(descAdicional, ".*\^([0-9.]+)", extraMults)
    if(extraMults1){
        MsgBox, Multiplying price by extra multiplier of %extraMults1%.
        mult := mult * extraMults1
    }
    RegExMatch(descAdicional, ".*\¬([0-9.]+)", extraDivisions)
    if(extraDivisions1){
        MsgBox, Dividing price by extra divisor of %extraDivisions1%.
        mult := mult / extraDivisions1
    }
    if(InStr(descAdicional, "½IVA")){
        if(mult != 1){
            ;MsgBox, Dividiendo el IVA a la mitad.
            mult := mult / 1.21 * 1.105
        }
    }
    
    if(not SeleccionarPrecio98o99()){
        return false
    }
    
    Clipboard := RegExReplace(Clipboard, ",", ".") ;Reemplazar comas por puntos.
    Clipboard := RegExReplace(Clipboard, "\.(?![^.]+$)")  ;Quitar todos los puntos excepto el último.
    Clipboard := RegExReplace(Clipboard, "[^0-9.]") ;Eliminar todo excepto números y puntos.
    if(parseNoDecimals == true){
        Clipboard := RegExReplace(Clipboard, "[.]") ;Eliminar puntos también.
    }
    if(not IsNum(Clipboard)) {
        MsgBox, Clipboard is not a number.
        return
    }
    
    multiplied := (Clipboard * mult)
    
    ControlGetText, oldPrice, %campoPrecioActual%, %ventanaPrecios%
    if(multiplied * 800 < oldPrice){
        multiplied := multiplied * 1000
    }
    
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.
    
    
    percent := (100*multiplied/oldPrice)-100
    percent := Round(percent, 1)
    if(percent < -15 or percent > 20){
        MsgBox, 305, , Diferencia de %percent%`%, continuar? ;1+48+256
        IfMsgBox, Cancel
        {
            Send, {Esc}
            Sleep, 150
            Send, {F10}
            return 0
        }
    }
    
    itemID := GetCodigoVentanaPrecios()
    LogPriceChange(itemID, oldPrice, multiplied, mult)
    
    ControlSetText, %campoPrecioActual%, %multiplied%, %ventanaPrecios%
    Send, {F10}
    Sleep, 150
    Send, {F10}
    return 1
}

MultiplicarPrecio98o99(mult:=0){
    if(mult == 0){
        explanation := "Ingrese el porcentaje a añadir o restar con el siguiente formato:`n1.21 o 21% para +21`%`n0.8 o -20% para -20`%"
        InputBox, multInput, Porcentaje, %explanation%,,,,,,,,%multiplicadorExtra%
        if ErrorLevel
            return ;Cancel
        mult := ParsePercent(multInput)
        if(!IsNum(mult) or mult == 0) {
            MsgBox, No se ingresó un número. (%mult%)
            return
        }
        multiplicadorExtra := multInput
    }
    
    if(not SeleccionarPrecio98o99()){
        return false
    }
    
    itemID := GetCodigoVentanaPrecios()
    ControlGetText, oldPrice, %campoPrecioActual%, %ventanaPrecios%
    multiplied := oldPrice*mult
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.
    
    LogPriceChange(itemID, oldPrice, multiplied, mult)
    
    ControlSetText, %campoPrecioActual%, %multiplied%, %ventanaPrecios%
    Send, {F10}
    Sleep, 150
    Send, {F10}
}
;}

;{ Navegación
SincronizarArticulosPrecio(){
    if(not WinExist(ventanaPrecios)){
        MsgBox No existe %ventanaPrecios%.
        return
    }
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }
    
    if(WinExist(ventanaBuscar)){
        CerrarVentanaBuscar()
    }
    
    if(SincronizadosArticulosPrecio()){
        return
    }

    CodigoArticulo := GetCodigoVentanaArticulos()
    
    GoToVentanaPrecios(CodigoArticulo)
}

LazySincronizarArticulosPrecio(){ ;Intento rudimentario para arreglar la desincronización
    if(WinExist(ventanaArticulos) and WinExist(ventanaPrecios)){
        WinWait, %ventanaArticulos%
        WinWait, %ventanaPrecios%
        CodigoArticulos := GetCodigoVentanaArticulos()
        CodigoArticulos := RegExReplace(CodigoArticulos, "[^0-9|\-|.]") ;Sólo numeros.
        CodigoPrecios := GetCodigoVentanaPrecios()
        CodigoPrecios := RegExReplace(CodigoPrecios, "[^0-9|\-|.]") ;Sólo numeros.
        if(CodigoArticulos > CodigoPrecios){
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Siguiente
        }
        else if(CodigoArticulos < CodigoPrecios){
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Siguiente
        }
    }
}

SincronizadosArticulosPrecio(){
    if(not WinExist(ventanaPrecios)){
        MsgBox SincronizadosArticulosPrecio - No existe %ventanaPrecios%.
        return false
    }
    if(not WinExist(ventanaArticulos)){
        MsgBox SincronizadosArticulosPrecio - No existe %ventanaArticulos%.
        return false
    }
    
    CodigoArticulos := GetCodigoVentanaArticulos()
    CodigoPrecios := GetCodigoVentanaPrecios()
    
    if(CodigoArticulos == CodigoPrecios){
        return true
    }
    else{
        return false
    }
}

ProximoArticulo(){ 
    next =
    if(enableCodeArray and CodeArray.Length() > 0 and WinExist(ventanaArticulos) and WinExist(ventanaPrecios)){
        if(GetCodigoVentanaArticulos() == CodeArray[CodeArray.MaxIndex()]){
            MsgBox Already at last element in CodeArray.
            return
        }
        next := NextFromArray(CodeArray, GetCodigoVentanaArticulos())
        if(RegExReplace(next, "[^0-9]", "") == RegExReplace(GetCodigoVentanaArticulos(), "[^0-9]", "") + 1){ ;numbers only
            next = ;just go forward normally
        }
    }
    
    if(next){
        GoToAll(next)
    }
    else{
        if(WinExist(ventanaArticulos)){
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Siguiente
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Siguiente
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Siguiente
        }
        if(WinExist(ventanaPrecios)){
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Siguiente
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Siguiente
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Siguiente
        }
        ;LazySincronizarArticulosPrecio()
    }   
}

AnteriorArticulo(){
    prev =
    if(enableCodeArray and CodeArray.Length() > 0 and WinExist(ventanaArticulos) and WinExist(ventanaPrecios)){
        if(GetCodigoVentanaArticulos() == CodeArray[1]){
            MsgBox Already at first element in CodeArray.
            return
        }
        prev := PrevFromArray(CodeArray, GetCodigoVentanaArticulos())
        if (RegExReplace(prev, "[^0-9]", "") == RegExReplace(GetCodigoVentanaArticulos(), "[^0-9]", "") - 1){ ;numbers only
            prev = ;just go backwards normally
        }
    }
    
    if(prev){
        GoToAll(prev)
    }
    else{
        if(WinExist(ventanaArticulos)){
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Anterior
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Anterior
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Anterior
        }
        if(WinExist(ventanaPrecios)){
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Anterior
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Anterior
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Anterior
        }
        ;LazySincronizarArticulosPrecio()
    }   
    

}

GoToAll(codigo){
    WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
    WinMenuSelectItem, %ventanaPrecios%, , Buscar, Por Clave
    
    WinWait, %ventanaBuscar_Articulos%,,5
    ControlFocus, %campoContenido_Buscar%, %ventanaBuscar_Articulos% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    Control, EditPaste, %codigo%, %campoContenido_Buscar%, %ventanaBuscar_Articulos%
    
    WinWait, %ventanaBuscar_Precios%,,5
    ControlFocus, %campoContenido_Buscar%, %ventanaBuscar_Precios% ;Si no hacemos focus, Tango no detecta que hicimos algún cambio.
    Control, EditPaste, %codigo%, %campoContenido_Buscar%, %ventanaBuscar_Precios%
    
    WinWait, %ventanaBuscar_Articulos%,,5
    ControlSend,, {Enter 2}, %ventanaBuscar_Articulos%
    WinWait, %ventanaBuscar_Precios%,,5
    ControlSend,, {Enter 2}, %ventanaBuscar_Precios%
}
;}

;{ Ventana Buscar
BuscarPorPortapapel(){
    if WinExist(ventanaCalc){
        WinActivate, %ventanaCalc%
        WinWait, %ventanaCalc%
        if WinExist(ventanaCalc_Buscar){
            WinActivate, Find & Replace
            WinWait, Find & Replace
            Send, !s ;Alt+S: Search For
        }
        else{
            Send, ^f ;Ctrl+F: Buscar
        }
        WinWait, Find & Replace
        Send, ^v{Enter} ;Ctrl+V+Enter
        WinWait, %ventanaCalc%
        if(PicExists("Images/OpenOfficeCalc/EndOf.png")){ ;Damn you, OpenOffice.
            Send, {Enter}
            WinWait, %ventanaCalc%
        }
        if(PicExists("Images/OpenOfficeCalc/NotFound.png")){ ;Damn you, OpenOffice.
            Send, {Enter}
            OnUnsuccessfulSearch()
            return 0
        }
        else{
            OnSuccessfulSearch()
            return 1
        }
    }
    else if WinExist(ventanaAdobeReader_Buscar){
        WinActivate, %ventanaAdobeReader_Buscar%
        WinWait, %ventanaAdobeReader_Buscar%
        ControlClick, %ventanaAdobeReader_BuscarOK%, %ventanaAdobeReader_Buscar%
        WinWait, %ventanaAdobeReader_Buscar%
        ControlFocus, %ventanaAdobeReader_Buscar_Input%, %ventanaAdobeReader_Buscar%
        Send, ^v{Enter} ;Ctrl+V+Enter
        WinWait, %ventanaAdobeReader_Buscar%
        ControlGetText, resultsText, %ventanaAbodeReader_Buscar_Matches%, %ventanaAdobeReader_Buscar%
        if(InStr(resultsText, "0 doc")){
            OnUnsuccessfulSearch()
            return 0
        }
        else{
            OnSuccessfulSearch()
            return 1
        }
        
    }
    else if WinExist(ventanaAdobeReader){
        WinActivate, %ventanaAdobeReader%
        WinWait, %ventanaAdobeReader%
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
        return 1
    }
}

OnUnsuccessfulSearch(){
}
OnSuccessfulSearch(){
    WinActivate, %ventanaCalc_Main%
    Send, %PostSearchString%
}

GetSearchTypeVentanaBuscar(){
    ControlGetText, searchType, %buscarEn%, %ventanaBuscar%
    return searchType
}
SetSearchTypeVentanaBuscar(searchType){
    Control, Choose, %searchType%, %buscarEn%, %ventanaBuscar%
}

CerrarVentanaBuscar(){
    if(not WinExist(ventanaBuscar)){
        MsgBox No existe %ventanaBuscar%.
        return
    }
    WinActivate, %ventanaBuscar%
    
    Control, Check, , %checkboxFiltrar%, %ventanaBuscar%
    Control, Uncheck, , %checkboxIncremental%, %ventanaBuscar%
    
    ControlGetText, CodigoIngresado, %campoContenido_Buscar%, %ventanaBuscar%
    if(CodigoIngresado == ""){
        Send, {Esc}
        WinWaitClose, %ventanaBuscar%
        return
    }
    
    Send, {Enter}
    If(WinExist(ventanaBuscar)){ ;Puede que ya hayamos apretado Enter nosotros.
        Send, {Enter}
    }
    ;WinWaitClose, %ventanaBuscar%
}
;}

;{ Logging
LogPriceChange(itemID := "", oldPrice := "", newPrice = "", mult := ""){
    percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    finalText = %itemID%: %percent%`% (x%mult%, %oldPrice% -> %newPrice%)
    if(codeArray){
        prog := ObjIndexOf(codeArray, itemID)
        length := codeArray.Length()
        if(prog){
            finalText = %finalText% - %prog%/%length% ;concatenation
        }
    }
    finalText = %finalText%`r`n ;concatenation
    LogSend(finalText)

}

LogArticleDeletion(itemID := "", oldDesc := ""){
    finalText = Eliminación: %itemID% (%oldDesc%)`r`n
    LogSend(finalText)
}

LogDescChange(itemID := "", oldDesc := "", newDesc = ""){
    finalText = %itemID%: %oldDesc% -> %newDesc%`r`n
    LogSend(finalText)
}

LogSend(finalText := ""){
    if(not WinExist(ventanaNotepad)){
        prev := WinActive("A")
        Run, Notepad
        WinWait, %ventanaNotepad%
        WinActivate, ahk_id %prev%
    }
    ControlSend,,^{End}, %ventanaNotepad% ;Ctrl+End: Go to end of document
    Control, EditPaste, %finalText%, , %ventanaNotepad%
}
;}

;{ Opciones
Menu, Tray, Add  ; Add a separator line.

Menu, Tray, Add, No Decimals, toggleNoDecimals
toggleNoDecimals(){
    if(parseNoDecimals == true){
        Menu, Tray, Uncheck, No Decimals
        parseNoDecimals := false
    }
    else{
        Menu, Tray, Check, No Decimals
        parseNoDecimals := true
    }
}

Menu, Tray, Add  ; Add a separator line.

Menu, Tray, Add, AdHocmode, toggleAdHocmode
toggleAdHocMode(){
    if(AdHocMode == true){
        Menu, Tray, Uncheck, AdHocmode
        AdHocMode := false
    }
    else{
        Menu, Tray, Check, AdHocmode
        AdHocMode := true
    }
}

Menu, Tray, Add  ; Add a separator line.

Menu, Tray, Add, Post-Search Commands..., setPostSearchString
setPostSearchString(){
    instructions := "Write out a set of instructions to send after a successful search.`rEach instruction must be between curly brackets, such as: {Right}`rAdd a number after your instruction to make it repeat that many times, for example: {Right 3}.`rSyntax is the same as AutoHotKey's Send command."
    InputBox, tempString, Post-Search Commands, %instructions%, , , , , , , , {}
    if(ErrorLevel or tempString == "" or tempString == "{}"){
        return
    }
    PostSearchString := tempString
    MsgBox, %PostSearchString%
}

Menu, Tray, Add  ; Add a separator line.

Menu, Tray, Add, Import CodeArray..., importCodeArray
importCodeArray(){
    FileSelectFile, archivo, 3, %A_Desktop%, Select a Tango-exported list...
    if(archivo = ""){
        return
    }
    
    codeArray := []
    Loop, read, %archivo%
    {
        Loop, parse, A_LoopReadLine, %A_Tab%
        {
            if(RegExMatch(A_LoopField, "^[\d*]+", match)){
                codeArray.Push(match)
            }
        }
    }
    
    array2text := StrJoin(CodeArray, ", ")
    MsgBox, 305, CodeArray Import, Import result:`r%array2text%`r`rGo to first entry? ;1+48+256
    IfMsgBox, OK
    {
        GoToAll(CodeArray[1])
    }
}

Menu, Tray, Add, CodeArray To Clipboard, copyCodeArray
copyCodeArray(){
    Clipboard := StrJoin(CodeArray, ", ")
    MsgBox, Copied:`n%Clipboard%
}

Menu, Tray, Add, Disable CodeArray, toggleCodeArray
toggleCodeArray(){
    if(enableCodeArray == true){
        Menu, Tray, Check, Disable CodeArray
        enableCodeArray := false
    }
    else{
        Menu, Tray, Uncheck, Disable CodeArray
        enableCodeArray := true
    }
}

Menu, Tray, Add  ; Add a separator line.
;{
Menu, searchTypeMenu, Add, %search_Default%, setSearchDefault, Radio
setSearchDefault(){
    setSearchType(search_Default)
}
Menu, searchTypeMenu, Add, %search_Exact%, setSearchExact, Radio
setSearchExact(){
    setSearchType(search_Exact)
}
Menu, searchTypeMenu, Add, %search_Start%, setSearchStart, Radio
setSearchStart(){
    setSearchType(search_Start)
}
Menu, searchTypeMenu, Add, %search_End%, setSearchEnd, Radio
setSearchEnd(){
    setSearchType(search_End)
}
Menu, searchTypeMenu, Add, %search_RemoveLastWord%, setSearchRemoveLastWord, Radio
setSearchRemoveLastWord(){
    setSearchType(search_RemoveLastWord)
}
Menu, searchTypeMenu, Add, %search_LongestNumber%, setSearchLongestNumber, Radio
setSearchLongestNumber(){
    setSearchType(search_LongestNumber)
}
Menu, searchTypeMenu, Add, %search_Fabrimport%, setSearchFabrimport, Radio
setSearchFabrimport(){
    setSearchType(search_Fabrimport)
}
Menu, searchTypeMenu, Add, %search_Faroluz%, setSearchFaroluz, Radio
setSearchFaroluz(){
    setSearchType(search_Faroluz)
}
Menu, searchTypeMenu, Add, %search_Ferrolux%, setSearchFerrolux, Radio
setSearchFerrolux(){
    setSearchType(search_Ferrolux)
}
Menu, searchTypeMenu, Add, %search_Solnic%, setSearchSolnic, Radio
setSearchSolnic(){
    setSearchType(search_Solnic)
}
;}
setSearchType(search_Default)

setSearchType(type){
    searchType := type
    
    Menu, searchTypeMenu, Uncheck, %search_Default%
    Menu, searchTypeMenu, Uncheck, %search_Exact%
    Menu, searchTypeMenu, Uncheck, %search_Start%
    Menu, searchTypeMenu, Uncheck, %search_End%
    Menu, searchTypeMenu, Uncheck, %search_RemoveLastWord%
    Menu, searchTypeMenu, Uncheck, %search_LongestNumber%
    Menu, searchTypeMenu, Uncheck, %search_Fabrimport%
    Menu, searchTypeMenu, Uncheck, %search_Faroluz%
    Menu, searchTypeMenu, Uncheck, %search_Ferrolux%
    Menu, searchTypeMenu, Uncheck, %search_Solnic%
    Menu, searchTypeMenu, Check, %type%
    
}

; Create a submenu in the first menu (a right-arrow indicator). When the user selects it, the second menu is displayed.
Menu, Tray, Add, Search Type, :searchTypeMenu

Menu, Tray, Add  ; Add a separator line.
Menu, Tray, Add, Exit, Exit
;}

;{ Misc
EstilizarVentanas(Activar := 1){
    if(Activar == 1){
        
        if(WinExist(ventanaArticulos_OKchild)){
            WinMove, %ventanaArticulos_OKchild%, , 370, 276
            return
        }
        
        if(WinExist(ventanaIngresosStock) and WinExist(ventanaEgresosStock)){
            WinMove, %ventanaIngresosStock%, , 255, 196
            WinActivate, %ventanaIngresosStock%
            WinMove, %ventanaEgresosStock%, , 255+572, 196
            WinActivate, %ventanaEgresosStock%
            return
        }
        
        WinSet, AlwaysOnTop, On, %ventanaArticulos%
        WinSet, Region, 0-0 W572 H222, %ventanaArticulos% ;Máscara de 572x222 empezando en 0,0
        WinMove, %ventanaArticulos%, , 1028, 26, 572
        WinGetPos, X, Y, W, H, %ventanaArticulos%
        
        WinSet, AlwaysOnTop, On, %ventanaPrecios%

        WinMove, %ventanaPrecios%, , X, Y+222, W, H
        WinGetPos, X, Y, W, H, %ventanaPrecios%
        
        if(WinExist(ventanaNotepad)){
            WinSet, Region, 0-0 W572 H398, %ventanaPrecios% ;Máscara de 572x425 empezando en 0,0
            
            WinSet, Region, 0-0 W999 H999, %ventanaNotepad% ;Literalmente sólo para que tenga los 3 pixeles negros feos
            WinSet, AlwaysOnTop, On, %ventanaNotepad%
            WinMove, %ventanaNotepad%, , X, Y+398 , W, 125
        }
        

    }
    else{
        WinSet, AlwaysOnTop, Off, %ventanaArticulos%
        WinSet, Region, , %ventanaArticulos%
        
        WinSet, AlwaysOnTop, Off, %ventanaPrecios%
        WinSet, Region, , %ventanaPrecios%
        
        WinSet, AlwaysOnTop, Off, %ventanaNotepad%
        WinSet, Region, , %ventanaNotepad%
    }
}

ParsePercent(input){
    if(InStr(input, "%")){
        input := RegExReplace(input, "[^0-9|\-|.]") ;Sólo numeros.
        return (100+input)/100
    }
    else{
        if(not IsNum(input)){
            return 0
        }
        return input
    }
}

AllRegexMatches(haystack, needle){
    Pos := 1
    Matches := []
    M := ""
    while(Pos := RegExMatch(haystack, needle, M, Pos + StrLen(M)))
    {
        Matches.Push(M)
    }
    return Matches
}

NextFromArray(arr, num){
    nextnum =
    for index, nextnum in arr{
        if(nextnum > num){
            return nextnum
        }
    }
}

PrevFromArray(arr, num){
    prevnum =
    for index, nextnum in arr{
        if(nextnum >= num){
            return prevnum
        }
        prevnum := nextnum
    }
}

ObjIndexOf(obj, item, case_sensitive:=false)
{
	for i, val in obj {
		if (case_sensitive ? (val == item) : (val = item))
			return i
	}
}

StrJoin(arr, del) {
    Result := ""
    for each, val in arr
        Result .= val del
    return RTrim(Result, del)
}

IsNum(str) { ;Fuck AHK.
	if str is number
		return true
	return false
}

IsAlwaysOnTop(Window) {
    WinGet, Estilo, ExStyle, %Window%
    Return (Estilo & 0x8) ; 0x8 is WS_EX_TOPMOST.
}
;}

;{ AdHoc
AdHoc(mult){
    if(not SincronizadosArticulosPrecio()){
        SincronizarArticulosPrecio()
        return
    }
    WinActivate, %ventanaCalc_Main%
    Send {Ctrl Down}c{Ctrl Up}
    if(not PegarPrecio98o99(mult)){
        return
    }
    ProximoArticulo()
    temp_medida := GetUnidadMedidaVentas()
    if(!temp_medida){
        return
    }
    Clipboard := temp_medida
    Sleep, 100
    BuscarPorPortapapel()
    WinActivate, %ventanaCalc_Main%
}
;}

;{ AUTOEXEC
if(not WinExist(ventanaNotepad)){
    Run, Notepad
}
return

Exit:
ExitApp
return
;}

;{ Keybinds
Launch_Media::
;EliminacionArticulo()
;PegarUnidadMedidaVentas()
MassActualizarDesc()
;Msgbox, Testing...
return

^Launch_Media::
ListLines 
return

Volume_Up::
EstilizarVentanas(1)
return

Volume_Down::
EstilizarVentanas(0)
return

Volume_Mute::
IngresarMultiplicadoresPrecio()
return

^Volume_Mute::
MultiplicarPrecio98o99()
return

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
medida := GetUnidadMedidaVentas()
if(!medida){
    return
}
Clipboard := medida
Sleep, 100
BuscarPorPortapapel()
return

Browser_Search::
medida := GetUnidadMedidaVentas()
if(!medida){
    return
}
Clipboard := medida
;CorregirUnidadMedidaVentas("Ferrolux")
Sleep,100
BuscarPorPortapapel()
return

Browser_Home::
if(AdHocMode){
    AdHoc(multiplicadorPrecio1)
}
else{
    PegarPrecio98o99(multiplicadorPrecio1)
}
return

^Browser_Home::
if(AdHocMode){
    AdHoc(multiplicadorPrecio2)
}
else{
    PegarPrecio98o99(multiplicadorPrecio2)
}
return
;}