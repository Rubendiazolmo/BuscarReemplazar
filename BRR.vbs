' Buscar y reemplazar en nombres de archivos de forma recursiva
On error resume Next
' Declaración de variables.
Dim ObjFSO, oShell, f, objFile, file
Dim Buscar, Reemplazar, NuevoNombre, NombreArchivoArray, NombreArchivo, NombreArchivoEncontrado, NuevoNombreArchivo

' Compruebo los parámetros de entrada del script
if WScript.Arguments.Count <> 2 then ' No tiene 2 parámetros de entrada

  ' Solicitar al usuario introducir los datos necesarios.
  Buscar = Inputbox("Buscar: ")
  ' Si se cancela termino la ejecución
  if IsEmpty(Buscar) then
    WScript.Quit
  end if

  Reemplazar = Inputbox("Reemplazar: ")
  ' Si se cancela termino la ejecución
  if IsEmpty(Reemplazar) then
    WScript.Quit
  end if

else ' Tiene 2 parámetros de entrada

  ' Leer parámetros de entrda del script si se ejecuta por consola introduciendo 2 parámetros de entrada
  Buscar    = Wscript.Arguments(0)
  Reemplazar = Wscript.Arguments(1)

end if

' Crear objeto para ejecutar comandos de sistema operativo
Set oShell = CreateObject ("WScript.Shell")
' Guardo los archivos encontrados en el archivo Archivos. Espero a que el comando se ejecute para seguir con el script
oShell.run "cmd.exe /C dir *" & Buscar & "* /b/s/a-d> Archivos",0,1 

' Creo objeto para poder trabajar con archivos
Set ObjFSO = CreateObject("Scripting.FileSystemObject") 
Set f      = ObjFSO.OpenTextFile("Archivos", 1, False)

' Iterara por cada línea del archivo Archivos (Por cada archivo que contenga el patrón buscado)  
Do Until f.AtEndOfStream

  file = f.ReadLine
  ' Reemplazo el patron buscado por el patron a reemplazar en la ruta del archivo

  NombreArchivoArray = Split(file, "\")
  NombreArchivoEncontrado = NombreArchivoArray(UBound(NombreArchivoArray))
  NuevoNombreArchivo = Replace(NombreArchivoEncontrado, Buscar, Reemplazar)
  
  NuevoNombre = Replace(file, NombreArchivoEncontrado, NuevoNombreArchivo)

   ' Creo y relleno el Txt
  ObjFSO.MoveFile file, NuevoNombre

loop

' Cierro el archivo
f.close

' Compruebo como se ha ejecutado el script, si no tiene parámetros de entrada (entiendo que ha sido por doble click)
' informo de que los txts se han generado con éxito.
if WScript.Arguments.Count <> 2 then


end if

' Borro el archivo Archivos
oShell.run "cmd.exe /C del Archivos",0,1
