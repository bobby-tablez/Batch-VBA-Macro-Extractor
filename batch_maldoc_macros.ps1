Write-Host "========================================================"
Write-Host "BATCH VB MACRO EXTRACTOR`nRequires oletools to be installed:`nGet it here: https://github.com/decalage2/oletools/";
Write-Host "========================================================`n`n"


$dir = Read-Host -Prompt "Enter maldoc directory (no trailing slash)";
Write-Host "Directory set to '$dir'";

$confirmExt = Read-Host -Prompt "Do we need to append extensions? [y/n]";
If($confirmExt -eq 'y') {
    $extension = Read-Host "Set Extension to: (Example: doc, xls, docm)"
    Get-ChildItem -Path "$dir" -File |Rename-Item -NewName { $PSItem.Name + "." + $extension }

} 
Elseif($confirmExt -eq 'n') {
    $extension = Read-Host -Prompt "Enter document file extension(s)";
    Write-Host "Extension set to '$extension'";
}
Else {
    Write-Warning "Expected [y/n] EXITING";
    Exit
}

Write-Host "Creating directory for extracted VB macros...";
New-Item -Path $dir'\extracted_vbs' -ItemType Directory -ErrorAction SilentlyContinue -ErrorVariable $DirError;
If ($DirError) { Write-Warning -Message "[WARNING] Directory create error... Directory likely exists or permissions error"; }

Write-Host "`nRunning against '$extension' files...";

Get-ChildItem $dir -Filter "*.$extension" | #get files
ForEach-Object{
    Try {
        $fileName = $_.Name;
        Write-Progress "Extracting VB macros from $fileName";
        olevba.exe -c $dir"\"$fileName >> $dir"\extracted_vbs\"$fileName".vbs" 2>&1;
    }
    Catch { Write-Warning "[WARNING] Possible error in $fileName"; }  
}
