;Richard Scheglmann - unsere-schule.org
#SingleInstance Force
#NoEnv
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

#Include %A_ScriptDir%\lib\GuiButtonIcon.ahk

Gui Font, s9, Segoe UI
Gui Add, Edit, vTextAreaLeft x30 y124 w300 h300 +Multi +HScroll  , ...
Gui Add, Edit, vTextAreaRight x377 y123 w300 h300 +Multi +HScroll  , ...
Gui Add, Button, gButtonUmbenennen x450 y444 w170 h23, Dateien umbenennen
Gui Add, Button, gButtonNeuLaden x87 y441 w170 h23, Dateien neu laden
Gui Add, Edit, vTextDateityp x30 y63 w74 h21
Gui Add, Edit, vTextSuchen x117 y63 w120 h21
Gui Add, Edit, vTextErsetzen x243 y63 w120 h21
Gui Add, Button, gButtonErsetzen x627 y45 w54 h33, GO
Gui Add, Button, gButtonDateinamenKopieren x341 y125 w25 h57, >
Gui Add, CheckBox, vCheckboxRegEx x298 y39 w67 h23, RegEx
Gui Add, Link, x305 y444 w120 h23, <a href="https://unsere-schule.org/programmieren/autohotkey/programme/einfacher-dateiumbenenner/">Hilfe und Anleitung</a>
Gui Add, Button, hWndhInfoBtn vInfoIcon gButtonOpenCurrentFolder x26 y3 w23 h23
Gui Add, Button, gButtonOrderWaehlen x50 y3 w120 h23, Ordner auswählen
Gui Add, Text, vPfadZumOrdner x175 y3 w509 h23 +0x200  , ...
Gui Add, GroupBox, x23 y98 w313 h335, aktuelle Dateinamen
Gui Add, GroupBox, x370 y98 w313 h335, neue Dateinamen
Gui Add, Edit, vPraefix x422 y39 w74 h21
Gui Add, Edit, vSuffix x422 y63 w74 h21
Gui Add, Text, x376 y42 w43 h23, Präfix
Gui Add, Text, x376 y65 w43 h23, Suffix
Gui Add, GroupBox, x114 y29 w253 h58, Suchen und Ersetzen
Gui Add, GroupBox, x371 y29 w133 h58
Gui Add, GroupBox, x24 y29 w86 h58, Dateityp
Gui Add, GroupBox, x508 y29 w102 h58, neuer Dateityp
Gui Add, Edit, vTextDateitypNeu x522 y62 w74 h21

GuiButtonIcon(hInfoBtn, "shell32.dll", 4, "L1 T1 w20 h20")

;globale Variablen
ordnerPfad := A_ScriptDir
fileExtension := ""
COUNTER := 1

GuiControl, text, PfadZumOrdner, %ordnerPfad%
GuiControl, text, TextDateityp, %fileExtension%
sucheDateien(True)

Gui Show, w709 h477, DateiUmbenenner V1.0.0
Return


GuiEscape:
GuiClose:
    ExitApp

sucheDateien(rechtesTextfeldFuellen){
    global ordnerPfad
    entryText := ""
    
    GuiControlGet, aktuelleDateiendung, ,TextDateityp 
    filesInFolder := ordnerPfad . "\*" . aktuelleDateiendung
          
    Loop, %filesInFolder%
    {
        entryText .= A_LoopFileName . "`n"
    }
            
    Sort, entryText
    
    ;füge die gefundenen Dateien den Textfeldern hinzu
    GuiControl, text, TextAreaLeft, %entryText%
    if(rechtesTextfeldFuellen = True){
        GuiControl, text, TextAreaRight, %entryText%
    }
}
ButtonOpenCurrentFolder(){
    global ordnerPfad
    Run, %ordnerPfad%    
}
GetAktuellenPfadZumOrdner(){
    global ordnerPfad
    GuiControlGet, ordnerPfad, ,PfadZumOrdner
}
ButtonDateinamenKopieren(){
    GuiControlGet, aktuelleDateinamen, ,TextAreaLeft    
    GuiControl, text, TextAreaRight, %aktuelleDateinamen%
}

ButtonUmbenennen() {
    global ordnerPfad
    
    MsgBox 0x34, Achtung, Wollen Sie wirklich die Dateien umbenennen?

    IfMsgBox Yes, {
        
        GuiControlGet, aktuelleDateinamen, ,TextAreaLeft    
        currentFileNamesList := StrSplit(aktuelleDateinamen, "`n")
        
        GuiControlGet, neueDateinamen, ,TextAreaRight    
        newFileNamesList := StrSplit(neueDateinamen, "`n")
        
        for index, element in currentFileNamesList
        {        
            newFileName := newFileNamesList[index]        
            FileMove, %ordnerPfad%\%element%, %ordnerPfad%\%newFileName%
        }
        GuiControl, text, TextAreaRight, ...
        sucheDateien(True)
    
    } Else IfMsgBox No, {
        return
    }
}

ButtonNeuLaden(){
    sucheDateien(False)
}

ButtonOrderWaehlen(){
    global ordnerPfad
    FileSelectFolder, ordnerPfadSelect, ,2
    ;wenn kein Pfad ausgewählt wurde, ändere nichts
    if(ordnerPfadSelect != ""){
        ordnerPfad := ordnerPfadSelect
    }
    GuiControl, text, PfadZumOrdner, %ordnerPfad%
    
    ;befülle linkes Textfeld mit Dateinamen
    sucheDateien(True)
}

addCounter(leadingZeros){
    global COUNTER
    
    txtCounter := ""
    txtLeadingZeroes := ""
    Loop, %leadingZeros% {
        txtLeadingZeroes .= "0"
    }
    txtCounter := SubStr(txtLeadingZeroes, 1, StrLen(txtLeadingZeroes) - StrLen(COUNTER)) . COUNTER    
    Return txtCounter
}

updateFilename(currentFilename){
    global COUNTER
    GuiControlGet, txtSuchen, ,TextSuchen
    GuiControlGet, txtErsetzen, ,TextErsetzen    
    GuiControlGet, checkedRegEx, ,CheckboxRegEx   
    
    newFilename := ""
    ;1) Suchen und Ersetzen + RegEx
    if(checkedRegEx = 0){
        newFilename := StrReplace(currentFilename, txtSuchen, txtErsetzen)
    }else{
        newFilename := RegExReplace(currentFilename, txtSuchen, txtErsetzen)
    }
    
    ;2) Präfix hinzufügen
    newFilename := getPraefix() . newFilename
    
    ;3) Suffix hinzufügen
    newFilename := newFilename . getSuffix()
    
    ;Counter wird pro Dateiname erhöht
    COUNTER := COUNTER + 1
    
    ;MsgBox, %newFilename%
    Return newFilename
}

getPraefix(){
    GuiControlGet, currentPraefix, ,Praefix 
    newPraefix := ""
    ;sucht einen Code und escaped diesem mit Pipes
    currentPraefix := StrReplace(currentPraefix, "<", "|<")
    currentPraefix := StrReplace(currentPraefix, ">", ">|")
    
    ;teilt den Suffix-String an den Pipes
    praefixArray := StrSplit(currentPraefix, "|")
    for index, element in praefixArray{
        newPraefix .= checkCodes(element)
    }  
    Return newPraefix    
}

getSuffix(){
    GuiControlGet, currentSuffix, ,Suffix
    newSuffix := ""
    ;sucht einen Code und escaped diesem mit Pipes
    currentSuffix := StrReplace(currentSuffix, "<", "|<")
    currentSuffix := StrReplace(currentSuffix, ">", ">|")
    
    ;teilt den Suffix-String an den Pipes
    suffixArray := StrSplit(currentSuffix, "|")
    for index, element in suffixArray{
        newSuffix .= checkCodes(element)
    }       
    Return newSuffix    
}

arrayToString(array){
    newString := ""
    for index, element in array{
        newString .= element . ","
    }
    return newString    
}

checkCodes(stringToCheck){   
    newString := ""
    tempCode := ""
    ;extrahiert alles innerhalb der Klammern <...>
    RegExMatch(stringToCheck, "(?<=<)([\w\._-]+)(?=>)", tempCode)
    if(tempCode != ""){
        ;special code
        if(InStr(tempCode, "C")){ ;Counter
            tempCode := StrReplace(tempCode, "C", "")
            leadingZeros := tempCode
            newString := addCounter(leadingZeros)
        }else if(InStr(tempCode, "T")){ ;Time and Date
            tempCode := StrReplace(tempCode, "T", "")
            FormatTime, currentDate , YYYYMMDDHH24MISS, %tempCode%
            newString := currentDate
        }
    
    }
    ;entfernt vorhandenen Code <...>
    newString .= RegExReplace(stringToCheck, "(<)([\w\._-]+)(>)", "")
    return newString
}

ButtonErsetzen(){
    global fileExtension, COUNTER
    COUNTER := 1
    GuiControlGet, txtSuchen, ,TextSuchen
    GuiControlGet, txtErsetzen, ,TextErsetzen    
    GuiControlGet, dateinamen, ,TextAreaLeft
    GuiControlGet, checkedRegEx, ,CheckboxRegEx   
    GuiControlGet, currentPraefix, ,Praefix 
    GuiControlGet, currentSuffix, ,Suffix      
    GuiControlGet, currentDateiTypNeu, ,TextDateitypNeu 
        
    dateinamen := RTrim(dateinamen, "`n")
    
    geaenderteDateinamen := ""
    
    dateinamenArray := StrSplit(dateinamen, "`n")
    for index, element in dateinamenArray{
    
        if(currentDateiTypNeu != ""){
            fileExtension := currentDateiTypNeu
        }else{
            fileExtension := getFileExtension(element)
        }
        
        element := removeFileExtension(element)
        geaenderteDateinamen .= updateFilename(element) . fileExtension . "`n"
    }

    GuiControl, text, TextAreaRight, %geaenderteDateinamen%
}

removeFileExtension(fileName){
    filenameWithoutExtension := RegExReplace(fileName, "\.\w+$", "")   
    Return filenameWithoutExtension
}

getFileExtension(fileName){
    RegExMatch(fileName, "\.(\w+$)", fileExtension)
    Return fileExtension
}
