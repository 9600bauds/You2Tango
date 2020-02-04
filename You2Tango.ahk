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

global ventanaArticulos_OKchild := "ahk_class TFormChildTg"
global campoOKChild_1 := "TNumEditTg4" ;the long awaited comeback to OK boomer
global campoOKChild_2 := "TEdit4"

global ventanaPrecios := "ACTUALIZACION DE PRECIOS INDIVIDUAL POR ARTICULO"
global campoCodigoArt_Precios_ModoNoModificar := "TEdit2" ;EN MODO NO MODIFICAR
global campoCodigoArt_Precios_ModoModificar := "TEdit6" ;EN MODO MODIFICAR
global campoPrecioActual := "TNumEditTg1"

global ventanaBuscar := "ahk_class TFrmBuscar"
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
global PostSearchString := ""
;}

;{ Ventana Artículos - Helpers
GetUnidadMedidaVentas(searchAfter := true){
    if(not WinExist(ventanaArticulos)){
        MsgBox No existe %ventanaArticulos%.
        return
    }

    ControlGetText, unidadMedida, %campoMedidaVentas%, %ventanaArticulos%
    
    if(searchAfter){
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
    
    if(WinExist(ventanaBuscar)){
        CerrarVentanaBuscar()
    }
    
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
    arr = 1578,2134,3758,5638,6544,500,2886
    
    if(not WinExist(ventanaArticulos)){
        MsgBox MassActualizarDesc - No existe %ventanaArticulos%.
        return
    }
    
    if(not WinExist(ventanaBuscar)){
        WinMenuSelectItem, %ventanaArticulos%, , Buscar, Por Clave
        WinWait, %ventanaBuscar%
    }
    
    Loop, parse, arr, `,,
    {
        ActualizarDescripFecha("search")
        Sleep, 250
        Send, %A_LoopField%
    }
    ActualizarDescripFecha()
}

CorregirUnidadMedidaVentas(prov := ""){ ;Desvergonzadamente ad-hoc. Requiere argumento.
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
    
    initialMedidaVentas := GetUnidadMedidaVentas(false)
    Clipboard := initialMedidaVentas
    if(prov == "Ferrolux"){
        Clipboard := RegExReplace(Clipboard, "([a-zA-Z])([1-9])","$1-$2")
        Clipboard := RegExReplace(Clipboard, " ","")
    }
    if(Clipboard != initialMedidaVentas){
        CambiarCampoVentanaArticulos(campoMedidaVentas, Clipboard)
    }
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
        MsgBox, Ventana Artículos y ventana Precios no están actualizadas!
        return false
    }
    
    if(not SeleccionarPrecio98o99()){
        return false
    }
    mult := ParsePercent(mult)
    if(mult == 0){
        MsgBox, Multiplicador inválido! - PegarPrecio98o99
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
    multiplied = % Round(multiplied, 2) ;Tango sólo quiere 2 decimales.
    
    itemID := GetCodigoVentanaPrecios()
    ControlGetText, oldPrice, %campoPrecioActual%, %ventanaPrecios%
    
    percent := (100*multiplied/oldPrice)-100
    percent := Round(percent, 1)
    if(percent < -15 or percent > 20){
        MsgBox, 305, , Diferencia de %percent%`%, continuar? ;1+48+256
        IfMsgBox, Cancel
        {
            Send, {Esc}
            Sleep, 150
            Send, {F10}
            return
            return 0
        }
    }
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
        next := NextFromArray(CodeArray, GetCodigoVentanaArticulos())
        if (next == GetCodigoVentanaArticulos() + 1){
            next = ;just go forward normally
        }
    }
    
    if(next){
        GoToVentanaArticulos(next, buscarCodigo)
        SincronizarArticulosPrecio()
    }
    else{
        if(WinExist(ventanaArticulos)){
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Siguiente
        }
        if(WinExist(ventanaPrecios)){
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Siguiente
        }
        LazySincronizarArticulosPrecio()
    }   
}

AnteriorArticulo(){
    prev =
    if(enableCodeArray and CodeArray.Length() > 0 and WinExist(ventanaArticulos) and WinExist(ventanaPrecios)){
        prev := PrevFromArray(CodeArray, GetCodigoVentanaArticulos())
        if (prev == GetCodigoVentanaArticulos() - 1){
            prev = ;just go backwards normally
        }
    }
    
    if(prev){
        GoToVentanaArticulos(prev, buscarCodigo)
        SincronizarArticulosPrecio()
    }
    else{
        if(WinExist(ventanaArticulos)){
            WinMenuSelectItem, %ventanaArticulos%, , Buscar, Anterior
        }
        if(WinExist(ventanaPrecios)){
            WinMenuSelectItem, %ventanaPrecios%, , Buscar, Anterior
        }
        LazySincronizarArticulosPrecio()
    }   
    

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
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
        Sleep, 100
        if(PicExists("Images/OpenOfficeCalc/EndOf.png")){ ;Damn you, OpenOffice.
            Send, {Enter}
        }
        Sleep, 100
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
    if WinExist(ventanaAdobeReader){
        WinActivate, %ventanaAdobeReader%
        WinWait, %ventanaAdobeReader%
        Send, ^f ;Ctrl+F: Buscar
        Sleep, 100
        Send, ^v{Enter} ;Ctrl+V+Enter
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
    WinWaitClose, %ventanaBuscar%
}
;}

;{ Logging
LogPriceChange(itemID := "", oldPrice := "", newPrice = "", mult := ""){
    percent := (100*newPrice/oldPrice)-100
    percent := Round(percent, 1)
    finalText = %itemID%: %percent%`% (x%mult%, %oldPrice% -> %newPrice%)`r`n
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

Menu, Tray, Add, Post-Search Commands..., setPostSearchString
setPostSearchString(){
    instructions := "Write out a set of instructions to send after a successful search.`rEach instruction must be between curly brackets, such as: {Right}`rAdd a number after your instruction to make it repeat that many times.`rSyntax is the same as AutoHotKey's Send command."
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
    FileSelectFile, archivo, 3, %A_Desktop%, Select a .txt document..., Text Documents (*.txt)
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
    
    MsgBox, , Import result:, % StrJoin(CodeArray, ", ")
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

Menu, searchTypeMenu, Add, %search_Default%, setSearchDefault, Radio
Menu, searchTypeMenu, Add, %search_Exact%, setSearchExact, Radio
Menu, searchTypeMenu, Add, %search_Start%, setSearchStart, Radio
Menu, searchTypeMenu, Add, %search_End%, setSearchEnd, Radio
Menu, searchTypeMenu, Add, %search_RemoveLastWord%, setSearchRemoveLastWord, Radio
Menu, searchTypeMenu, Add, %search_LongestNumber%, setSearchLongestNumber, Radio
Menu, searchTypeMenu, Add, %search_Fabrimport%, setSearchFabrimport, Radio
Menu, searchTypeMenu, Add, %search_Faroluz%, setSearchFaroluz, Radio
Menu, searchTypeMenu, Add, %search_Ferrolux%, setSearchFerrolux, Radio
Menu, searchTypeMenu, Add, %search_Solnic%, setSearchSolnic, Radio
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

;{ AUTOEXEC
if(not WinExist(ventanaNotepad)){
    Run, Notepad
}
return
;}

;{ Opciones - post autoexec
toggleNoDecimals:
toggleNoDecimals()
return

setPostSearchString:
setPostSearchString()
return

toggleCodeArray:
toggleCodeArray()
return

importCodeArray:
importCodeArray()
return

setSearchDefault:
setSearchType(search_Default)
return

setSearchExact:
setSearchType(search_Exact)
return

setSearchStart:
setSearchType(search_Start)
return

setSearchEnd:
setSearchType(search_End)
return

setSearchRemoveLastWord:
setSearchType(search_RemoveLastWord)
return

setSearchLongestNumber:
setSearchType(search_LongestNumber)
return

setSearchFabrimport:
setSearchType(search_Fabrimport)
return

setSearchFaroluz:
setSearchType(search_Faroluz)
return

setSearchFerrolux:
setSearchType(search_Ferrolux)
return

setSearchSolnic:
setSearchType(search_Solnic)
return

Exit:
ExitApp
return
;}

;{ Keybinds
Launch_Media::
;EliminacionArticulo()
;PegarUnidadMedidaVentas()
Msgbox, Testing...
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
Clipboard := GetUnidadMedidaVentas()
Sleep, 100
BuscarPorPortapapel()
return

Browser_Search::
Clipboard := GetUnidadMedidaVentas()
;CorregirUnidadMedidaVentas("Ferrolux")
Sleep,100
BuscarPorPortapapel()
return

Browser_Home::
PegarPrecio98o99(multiplicadorPrecio1)
return

^Browser_Home::
PegarPrecio98o99(multiplicadorPrecio2)
return
;}