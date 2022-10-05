Set WshShell = WScript.CreateObject("WScript.Shell")
Set sistem = CreateObject("Scripting.FileSystemObject")

do
    do
        Dim message, title, default, value
            message = "Qual o caminho da pasta?"& vbCrLf &"caso deseje parar app digite pare"
            title = "Procurando caminho . . . ."
            default = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")

        caminho = InputBox(message, title, default+"\")
        If IsEmpty(caminho) Then
            MsgBox("preciso informar caminho")
        ElseIf  IsNull(caminho) Then
            MsgBox("preciso informar caminho")
        ElseIf caminho="" Then
            MsgBox("preciso informar caminho")
        ElseIf caminho="pare" Then
            WScript.Quit()
        Else
            MsgBox(caminho)
            Exit Do
        End If
    loop

    paths =  sistem.GetAbsolutePathName(caminho)
    set folders = sistem.GetFolder(paths)
    set sisFolders = folders.SubFolders

    For Each objFolder in sisFolders
        If objFolder.Name =".metadata" Then
            createobject("wscript.shell").popup objFolder.Name+" nao pode fazer pull", 2, "PULL DIARIO"
        ElseIf objFolder.Name =".git" Then
            createobject("wscript.shell").popup objFolder.Name+" nao pode fazer pull", 2, "PULL DIARIO"
        ElseIf objFolder.Name ="node_modules" Then
            createobject("wscript.shell").popup objFolder.Name+" nao pode fazer pull", 2, "PULL DIARIO"
        Else
            WshShell.Run "cmd.exe",0,False
                WScript.Sleep(1000)
            WshShell.SendKeys("cd ")
            WshShell.SendKeys(paths+"\"+objFolder.Name)
            Wshshell.sendkeys "{ENTER}"
                WScript.Sleep(1000)
            WshShell.SendKeys("git pull")
                WScript.Sleep(1000)
            Wshshell.sendkeys "{ENTER}"
                WScript.Sleep(8000)
            WshShell.Run("taskkill /im cmd.exe")
            createobject("wscript.shell").popup objFolder.Name+" finalizado pull", 2, "PULL DIARIO"
        End if
    Next

    MsgBox("todas as pastas foram atualizadas")
    Dim mess, tit, def
    tit = "PULL DIARIO"
    mess = "fazer em outra pasta?"& vbCrLf &"digite n para parar"& vbCrLf &"caso cancele ou digite nada a aplicacao para"
    def = "n"

    denovo = InputBox(mess, tit, def)
    If IsNull(denovo) Then
        WScript.Quit()
    ElseIf IsEmpty(denovo) Then
        WScript.Quit()
    ElseIf denovo="n" Then
        WScript.Quit()
    Else
        End If
loop
