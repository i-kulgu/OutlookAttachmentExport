#
# Voeg benodigde Assembly items toe
#
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-type -assembly “Microsoft.Office.Interop.Outlook”

#
# Te gebruiken variabelen binnen script
#

$outlook = New-Object -comobject outlook.application
$word = New-Object -ComObject Word.Application
$namespace = $outlook.GetNamespace("MAPI")
$prefix = Get-Date -Format yyyyMMdd_HHmmss_
$folder = $namespace.PickFolder()
$filepath = "C:\Temp\OutlookAttachments"
$archivepath = " " # Voer een padnaam in om te archiveren
$adobe = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"

#
# Pak alle bijlagen in het geselecteerde folder binnen Outlook
#

$folder.Items| foreach {
 $SendName = $_.SenderName
   $_.attachments|foreach {
    $a = $_.filename

    #
    # Als bijlage een bestandsextensie .doc(x) of .pdf heeft wordt het geexporteerd naar $filepath hierboven met het opgegeven $prefix
    #

    If ($a.Contains("doc") -or $a.Contains("pdf") ) {
    $_.saveasfile((Join-Path $filepath ($prefix + $a)))
   }
  }
}

#
# Pak alle bestanden in $filepath folder, als het een .doc(x) extensie heeft zet het om naar pdf.
#

Get-ChildItem -Path $filepath | ForEach-Object {
if ($_.FullName -like "*.doc?"){
    $document = $word.Documents.Open($_.FullName)
    $pdf_filename = "$($filepath)\$($_.BaseName).pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    $document.Close()
    }
}
$word.Quit()

#
# Verwijder vervolgens alle .doc(x) bestanden die in $filepath zitten
#

Get-ChildItem -Path $filepath -Filter *.doc? | Remove-Item

#
# Pak alle pdf bestanden in $filepath op en print het naar standaard printer
#

$pdfbestanden = Get-ChildItem -Path $filepath -Filter *.pdf 

foreach ($pdf in $pdfbestanden) {
    $arglist = '/S /T "' + $pdf.FullName + '"'
    Start-Process $adobe -ArgumentList $arglist 
	start-sleep -s 5
} 

#
# Adobe laat altijd het scherm open nadat je geprint hebt.
# Hiermee kijken we of er adobe processen lopen en zo ja, dan sluit die ze allemaal.
#

Get-Process -Name AcroRd32 | Stop-Process

#
# Controleer of het archiefpad bestaat, zo niet, maak het recursief aan.
# Kopieer vervolgens alle bestanden vanuit $filepath naar $archivepath
#

if (!(Test-Path $archivepath)){
    New-Item $archivepath -ItemType Directory
} else {
    Get-ChildItem -Path $filepath -Filter *.pdf | ForEach-Object { Move-Item -path $_.FullName -destination $archivepath }
}

#
# Als laatst verwijderen we alle pdf bestanden die in $filepath staan.
#

Get-ChildItem -Path $filepath | Remove-Item -Force -Confirm:$false


[System.Windows.MessageBox]::Show('Printen is afgerond. Vergeet niet je printjes op te halen.')
