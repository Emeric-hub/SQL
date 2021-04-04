strInitialDB = "master" 'default database
IdSrv = "01-SAGE-01"

'strSQLServer = "SQL" & IdSrv & "\SQLSERVER2008" 'type the name of SQL Server instance here.
strSQLServer = "127.0.0.1"
pathsvg = "CHEMIN a REMPLACER"
pathan = pathsvg & DatePart("yyyy",now,vbSunday,vbFirstFourDays)
pathsem = pathsvg & DatePart("yyyy",now,vbSunday,vbFirstFourDays) & "\" & DatePart("ww",now,vbSunday,vbFirstFourDays)
pathlocal = "CHEMIN_LOCAL_BACKUP"
pathlog = pathlocal & "\logs\"
ficlog = pathlog & "Backup_COMPL_" & DatePart("d",now,vbSunday,vbFirstFourDays) & ".log"
dim DateBDD

set Shell = WScript.CreateObject("WScript.Shell")

Ret1 = Shell.run("cmd /C net use " & pathsvg & " /user:compte_de_service mdp_du_compte",1,TRUE)

Set objFSODate = CreateObject("Scripting.FileSystemObject")
	objFSODate.DeleteFile "c:\UESScripts\Liste_bases.txt", True
Set LogScrDate = objFSODate.OpenTextFile ("c:\UESScripts\Liste_bases.txt", 8, True)
 
Set LogSVG = objFSODate.OpenTextFile ("C:\UESScripts\log\Sauve_Toutes_Bases_NAS.log", 8, True)
LogSVG.WriteLine(now & " - Suppression fichiers de " & pathlocal)

	Ret1 = Shell.run("cmd /C del " & pathlocal & "*.* /F/Q",1,TRUE)

LogSVG.WriteLine(now & " - Démarrage sauvegarde C:\UESScripts\Sauve_Toutes_Bases_NAS.vbs ")

Set objConn = CreateObject("ADODB.Connection")

objConn.Open "Provider=SQLOLEDB.1;Password=MDP_SQLSERVER;Persist Security Info=False;User ID=sa;Initial Catalog=" & strInitialDB & ";Data Source=" & strSQLServer

strQuery = "SELECT [name] FROM [sys].[databases] where [database_id] >  '4' and [state] = '0' order by [name]"
set objRS = objConn.execute(strQuery)

nbbases=0
while not objRS.EOF

	nbbases=nbbases+1
	Set objConnDate = CreateObject("ADODB.Connection")
		objConnDate.ConnectionTimeout = 30
		objConnDate.Open "Provider=SQLOLEDB.1;Password=MDP_SQLSERVER;Persist Security Info=False;User ID=sa;Database=" & objRS.fields(0).value & ";Data Source=" & strSQLServer
		Svg_Base(objRS.fields(0).value)
		LogScrDate.WriteLine(objRS.fields(0).value)
		objConnDate.close
		Set objConnDate = Nothing 
		set objDate = Nothing

	objRS.moveNext
wend

LogScrDate.Close

LogSVG.WriteLine(now & " - Fin sauvegarde C:\UESScripts\Sauve_Toutes_Bases_NAS.vbs ")
LogSVG.WriteLine(now & " - Demarrage transfert " & pathlocal & " vers " & pathsem)
LogSVG.Close

	Ret1 = Shell.run("cmd /C mkdir " & pathan,1,TRUE)
	Ret2 = Shell.run("cmd /C mkdir " & pathsem,1,TRUE)
	Ret0 = Shell.run("cmd /C ROBOCOPY " & pathlocal & " " & pathsem & "\ /MIR",1,True)

Set LogSVG = objFSODate.OpenTextFile ("C:\UESScripts\log\Sauve_Toutes_Bases_NAS.log", 8, True)
LogSVG.WriteLine(now & " - Fin transfert " & pathlocal & " vers " & pathsem)
LogSVG.WriteLine("-------------------------------------------------------------------------------")
LogSVG.Close

'------------------------------------------------------------------------------------------------------------------------------

sub Svg_Base(Nmbase)
	
  	    Ret3 = Shell.run("cmd /C echo " & now & " --------- TOTALE -------------------------------------------------------------------------------------------------- >> " & ficlog,1,True)
	    Re43 = Shell.run("cmd /C sqlcmd -S " & strSQLServer & " -U sa -P MDP_SQLSERVER -Q " & chr(34) & "BACKUP DATABASE ["& NmBase & "] TO DISK = '" & pathlocal & "\" & NmBase & ".BAK' WITH NOFORMAT,INIT,NAME = 'Sauvegarde base " & NmBase & "'" & chr(34) & " >> " & ficlog,1,true)
end sub