<<<<<<< HEAD
Option Explicit

function dateienFinden()
    Dim objFSO, ordner, Subfolders, unterordner, i, Files

    Set objFSO = createobject("Scripting.FileSystemObject")
    Set ordner = objFSO.getfolder("E:\Users\chris\Documents\data\data\")
    i = 0
    
    Set Subfolders = ordner.subfolders

    If Subfolders.Count > 0 Then
        For each unterordner in Subfolders
            'Bekomme Pfad von index.html
            Redim Preserve arrListe(i)
            arrListe(i) = unterordner.Path & "\index.html"
            i = i + 1
        next
    End If
    dateienFinden = arrListe
End Function

function dateiAuslesen (strPfad)
    Dim objFSO, objFile, objReadFile, i

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(strPfad)

    'If objFile.FileExists Then
        Set objReadFile = objFSO.OpenTextFile(strPfad, 1)
        i = 0

        Do Until objReadFile.AtEndOfStream
            Redim Preserve arrData(i)
            arrData(i) = objReadFile.ReadLine
            i = i + 1
        Loop
        objReadFile.Close
        dateiAuslesen = arrData
    'Else
      '  MsgBox "Empty File"
     '   dateiAuslesen = null
    'End If
End Function

function cleanTags(strInput)
    Dim objRegexp
    Set objRegexp = New RegExp
    objRegexp.Global = True
    objRegexp.IgnoreCase = True
    objRegexp.Multiline = False
    objRegexp.Pattern = "(<[^<]*>|;|&nbsp)"

    cleanTags = objRegexp.Replace(strInput, "")

End Function

function dateiAuswerten(arrInput)
    Dim suchterm(16), strTemp, strElement, strDatensatz, objRegexp, colMatches, match, j, strCleanTemp, temp, tem
    suchterm(0) = "(\d{2,3})(?:.{,5})?(?:dauer|played|spielzeit)(?:.{,5})(\d{2,3})?" 'Dauer
    suchterm(1) = "<td class=""vi-is1-solid vi-is1-tbll"">(.*)[^<]*</span></span></td>" 'preis
    suchterm(2) = "(\d{2,3})?(.{,5})?(beruf|(\w+(kunst|kunde)))(.{,5})?(\d{2,3})?" 'craft
    suchterm(3) = "dual((\s)?spec)?" 'dual
    suchterm(4) = "(lv(l)?|level|stufe|\d{2}er)\s(\d{2})" 'lvl
    suchterm(5) = "(kalt(wetter)?|epi(c|sch) (fliegen|reiten)|reit(en)?|flug|coldweather)" 'reiten
    suchterm(6) = "" 'classq done
    suchterm(7) = "\d{,7}?.{,4}(emblem(e)?\s(of\sconquest|der\sEroberung)).{,4}\d{,7}?" 'conquest
    suchterm(8) = "(\d{1,3}k|\d{1,7})?(ehre((n)?punkt(e)?)) (\d{1,3}k|\d{1,7})?" 'ehre
    suchterm(9) = "(.{20})(equip|item(lvl|level))(.{,6}\d{3}|.{20})" 'eq
    suchterm(10) = "" 'flugpkt done
    suchterm(11) = "(\d{1,7}|\d{1,4}k)?.{,6}(gold|geld|bar|vermögen).{,6}(\d{1,7}|\d{1,4}k)?" 'geld
    suchterm(12) = "" 'hero done
    suchterm(13) = "(mount)?.{,2}((\d{2,3}%)).{,2}(mount)?|(epi(c|sch) (fliegen|flugmount))" 'mount
    suchterm(14) = "ehrfürchtig" 'ruf
    suchterm(15) = "(\d{,7}|\d{1,3}k)?.{,4}(emblem(e)?\s(of|des)\striumph(s)?).{,4}(\d{,7}|\d{1,3}k)?" 'triumph
    suchterm(16) = "<div id=""vi-content"">.*<a href="".*"">zurück zur Startseite</a>" 'auktionstext

    For j = 0 To (UBound(arrInput))
        strTemp = strTemp & arrInput(j)
    Next

    strCleanTemp = cleanTags(strTemp)

    'MsgBox strTemp

    'Ein Suchterm nach dem Anderen!
    For j = 0 To (UBound(suchterm))

        If j = 1 Or j = 16 Then
            Set objRegexp = New RegExp
            objRegexp.Global = True
            objRegexp.IgnoreCase = True
            objRegexp.Multiline = True
            objRegexp.Pattern = suchterm(j)
    
            If j=4 Or j = 9 Then
                objRegexp.Global = False
            End If
    
            Set colMatches = objRegExp.Execute(strTemp)
            strElement = ""
            If ((colMatches.Count = 0) Or (j = 6) Or (j = 10) Or (j = 12))  Then
                strElement = "-"
            Else 
                For Each match In colMatches
                    strElement = strElement + cleanTags(match.Value)
                Next
            End If
            'Umkodierung einfügen
            'clean text einfügen
            'Element des Datensatzes hinzufügen
            strDatensatz = strDatensatz & strElement & ";"
        Else


            Set objRegexp = New RegExp
            objRegexp.Global = True
            objRegexp.IgnoreCase = True
            objRegexp.Multiline = True
            objRegexp.Pattern = suchterm(j)
    
            If j=4  Then
                objRegexp.Global = False
            End If
    
            Set colMatches = objRegExp.Execute(strCleanTemp)
            strElement = ""
            If ((colMatches.Count = 0) Or (j = 6) Or (j = 10) Or (j = 12))  Then
                strElement = "-"
            Else 
                For Each match In colMatches
                    strElement = strElement + cleanTags(match.Value)
                Next
            End If
            'Umkodierung einfügen
            'clean text einfügen
            'Element des Datensatzes hinzufügen
            strDatensatz = strDatensatz & strElement & ";"
        End If
    Next
    dateiAuswerten = strDatensatz
    'msgbox "daten abgefangen"
End Function

function datensatzSchreiben(strDatensatz)
    Dim fs, output
    set fs = createobject("Scripting.filesystemobject")
    set output = fs.opentextfile("e:\users\chris\desktop\output.txt",8,true,-2)
    output.writeline strDatensatz
    output.close
End Function

'---------------------------
'Beginn der Programmschleife
'vvvvvvvvvvvvvvvvvvvvvvvvvvvv
Dim strPfad, arrDateiContent, strDatensatz, arrDateiliste, i

arrDateiliste = dateienFinden()
For i=0 To (Ubound(arrDateiliste))
    strPfad = arrDateiliste(i)

    If (strPfad > "") Then
        'msgbox ("eins:" & i)

        arrDateiContent = dateiAuslesen(strPfad)
        'msgbox arrDateiContent(50)
        strDatensatz = dateiAuswerten(arrDateiContent)
        'msgbox strDatensatz
        datensatzSchreiben strDatensatz
    End If
=======
Option Explicit

function dateienFinden()
    Dim objFSO, ordner, Subfolders, unterordner, i, Files

    Set objFSO = createobject("Scripting.FileSystemObject")
    Set ordner = objFSO.getfolder("E:\Users\chris\Documents\data\data\")
    i = 0
    
    Set Subfolders = ordner.subfolders

    If Subfolders.Count > 0 Then
        For each unterordner in Subfolders
            'Bekomme Pfad von index.html
            Redim Preserve arrListe(i)
            arrListe(i) = unterordner.Path & "\index.html"
            i = i + 1
        next
    End If
    dateienFinden = arrListe
End Function

function dateiAuslesen (strPfad)
    Dim objFSO, objFile, objReadFile, i

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(strPfad)

    'If objFile.FileExists Then
        Set objReadFile = objFSO.OpenTextFile(strPfad, 1)
        i = 0

        Do Until objReadFile.AtEndOfStream
            Redim Preserve arrData(i)
            arrData(i) = objReadFile.ReadLine
            i = i + 1
        Loop
        objReadFile.Close
        dateiAuslesen = arrData
    'Else
      '  MsgBox "Empty File"
     '   dateiAuslesen = null
    'End If
End Function

function cleanTags(strInput)
    Dim objRegexp
    Set objRegexp = New RegExp
    objRegexp.Global = True
    objRegexp.IgnoreCase = True
    objRegexp.Multiline = False
    objRegexp.Pattern = "(<[^<]*>|;|&nbsp)"

    cleanTags = objRegexp.Replace(strInput, "")

End Function

function dateiAuswerten(arrInput)
    Dim suchterm(16), strTemp, strElement, strDatensatz, objRegexp, colMatches, match, j, strCleanTemp, temp, tem
    suchterm(0) = "(\d{2,3})(?:.{,5})?(?:dauer|played|spielzeit)(?:.{,5})(\d{2,3})?" 'Dauer
    suchterm(1) = "<td class=""vi-is1-solid vi-is1-tbll"">(.*)[^<]*</span></span></td>" 'preis
    suchterm(2) = "(\d{2,3})?(.{,5})?(beruf|(\w+(kunst|kunde)))(.{,5})?(\d{2,3})?" 'craft
    suchterm(3) = "dual((\s)?spec)?" 'dual
    suchterm(4) = "(lv(l)?|level|stufe|\d{2}er)\s(\d{2})" 'lvl
    suchterm(5) = "(kalt(wetter)?|epi(c|sch) (fliegen|reiten)|reit(en)?|flug|coldweather)" 'reiten
    suchterm(6) = "" 'classq done
    suchterm(7) = "\d{,7}?.{,4}(emblem(e)?\s(of\sconquest|der\sEroberung)).{,4}\d{,7}?" 'conquest
    suchterm(8) = "(\d{1,3}k|\d{1,7})?(ehre((n)?punkt(e)?)) (\d{1,3}k|\d{1,7})?" 'ehre
    suchterm(9) = "(.{20})(equip|item(lvl|level))(.{,6}\d{3}|.{20})" 'eq
    suchterm(10) = "" 'flugpkt done
    suchterm(11) = "(\d{1,7}|\d{1,4}k)?.{,6}(gold|geld|bar|vermögen).{,6}(\d{1,7}|\d{1,4}k)?" 'geld
    suchterm(12) = "" 'hero done
    suchterm(13) = "(mount)?.{,2}((\d{2,3}%)).{,2}(mount)?|(epi(c|sch) (fliegen|flugmount))" 'mount
    suchterm(14) = "ehrfürchtig" 'ruf
    suchterm(15) = "(\d{,7}|\d{1,3}k)?.{,4}(emblem(e)?\s(of|des)\striumph(s)?).{,4}(\d{,7}|\d{1,3}k)?" 'triumph
    suchterm(16) = "<div id=""vi-content"">.*<a href="".*"">zurück zur Startseite</a>" 'auktionstext

    For j = 0 To (UBound(arrInput))
        strTemp = strTemp & arrInput(j)
    Next

    strCleanTemp = cleanTags(strTemp)

    'MsgBox strTemp

    'Ein Suchterm nach dem Anderen!
    For j = 0 To (UBound(suchterm))

        If j = 1 Or j = 16 Then
            Set objRegexp = New RegExp
            objRegexp.Global = True
            objRegexp.IgnoreCase = True
            objRegexp.Multiline = True
            objRegexp.Pattern = suchterm(j)
    
            If j=4 Or j = 9 Then
                objRegexp.Global = False
            End If
    
            Set colMatches = objRegExp.Execute(strTemp)
            strElement = ""
            If ((colMatches.Count = 0) Or (j = 6) Or (j = 10) Or (j = 12))  Then
                strElement = "-"
            Else 
                For Each match In colMatches
                    strElement = strElement + cleanTags(match.Value)
                Next
            End If
            'Umkodierung einfügen
            'clean text einfügen
            'Element des Datensatzes hinzufügen
            strDatensatz = strDatensatz & strElement & ";"
        Else


            Set objRegexp = New RegExp
            objRegexp.Global = True
            objRegexp.IgnoreCase = True
            objRegexp.Multiline = True
            objRegexp.Pattern = suchterm(j)
    
            If j=4  Then
                objRegexp.Global = False
            End If
    
            Set colMatches = objRegExp.Execute(strCleanTemp)
            strElement = ""
            If ((colMatches.Count = 0) Or (j = 6) Or (j = 10) Or (j = 12))  Then
                strElement = "-"
            Else 
                For Each match In colMatches
                    strElement = strElement + cleanTags(match.Value)
                Next
            End If
            'Umkodierung einfügen
            'clean text einfügen
            'Element des Datensatzes hinzufügen
            strDatensatz = strDatensatz & strElement & ";"
        End If
    Next
    dateiAuswerten = strDatensatz
    'msgbox "daten abgefangen"
End Function

function datensatzSchreiben(strDatensatz)
    Dim fs, output
    set fs = createobject("Scripting.filesystemobject")
    set output = fs.opentextfile("e:\users\chris\desktop\output.txt",8,true,-2)
    output.writeline strDatensatz
    output.close
End Function

'---------------------------
'Beginn der Programmschleife
'vvvvvvvvvvvvvvvvvvvvvvvvvvvv
Dim strPfad, arrDateiContent, strDatensatz, arrDateiliste, i

arrDateiliste = dateienFinden()
For i=0 To (Ubound(arrDateiliste))
    strPfad = arrDateiliste(i)

    If (strPfad > "") Then
        'msgbox ("eins:" & i)

        arrDateiContent = dateiAuslesen(strPfad)
        'msgbox arrDateiContent(50)
        strDatensatz = dateiAuswerten(arrDateiContent)
        'msgbox strDatensatz
        datensatzSchreiben strDatensatz
    End If
>>>>>>> c_ebayssh/master
Next