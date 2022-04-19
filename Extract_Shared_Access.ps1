<#
Autor : SAUTRON Dimitri
Description : Extract accesses from a shared directory.
  Extraction of accesses on the directory either users or groups.
  If the access is managed by a group, an extraction of the group is provided.
  The output file will be in xlsx format.
Version : 1.2
Date : 04/10/2022
#>

######################################################################
Import-Module activedirectory

$date = Get-date -Format yyyy-MM-dd_HH-mm
Start-Transcript "C:\Tools\AD\Share_Access-$date.txt"

$File = "C:\Tools\AD\Export_Informatique.xlsx"
$Excel = New-Object -ComObject Excel.Application
$WorkBook = $Excel.Workbooks.Open($File)
$worksheet = $workbook.worksheets.Item("Informatique")

$y=1
$x=1

#set-location -path "D:\data\Services\Informatique"
$ShareACL = "\\Serveur\Fichier\Informatique\"
$repertoires_SOC = Get-ChildItem $ShareACL -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $true)} | Select-Object Name,FullName

Foreach ($rep in $repertoires_SOC)
{
    
    $Excel_Export = New-Object PSObject
    # Coupage de l'héritage du dossier SOCIAL en cours (en conservant les accès en cours)
    $rep_social = Get-Item $rep.FullName
    $acl = $rep_social|get-acl

        # Definir l'emplacement du chemin
        $worksheet.Cells.Item($y,$x) = $rep.FullName
        $acl = $acl.Access | Where-Object {($_.IdentityReference -notlike "*BUILTIN\Administrateurs*") -and ($_.IdentityReference -notlike "*Admins du domaine*")}
        Write-host "Chemin : "$rep.FullName -BackgroundColor Yellow -ForegroundColor blue
        Write-host "Repertoire : "$rep.Name -BackgroundColor Red -ForegroundColor black
        Foreach($acc in $acl)
    	    {
                $GroupAD = $acc.IdentityReference
                $Grp_AD = $GroupAD -replace [regex]::Escape('domain.local\'),('')
                Write-Host "ACL : "$Grp_AD -BackgroundColor Green -ForegroundColor black
                Write-Host "Type ACL : "$acc.FileSystemRights -BackgroundColor Green -ForegroundColor black
                # Definir l'emplacement du groupe
                $worksheet.Cells.Item($y+1,$x+1) = $Grp_AD
                $Real_FileSystemRights = $acc.FileSystemRights.ToString()
                # Definir l'emplacement du type ACL
                $worksheet.Cells.Item($y+1,$x+2) = $Real_FileSystemRights
                $CheckObject = Get-ADObject -Filter {(sAMAccountName -eq $Grp_AD)} -Properties Name,sAMAccountName,ObjectClass

                if($CheckObject.ObjectClass -eq "group")
                {
                    $lol = Get-ADGroup $CheckObject | Select-Object Name
                    $Member_GRP = Get-ADGroupMember -Identity $lol.Name | Select-Object Name
                    foreach($GrpMembers in $Member_GRP)
                    {
                        $y++
                        $worksheet.Cells.Item($y+1,$x+2) = $GrpMembers.Name
                        #write-host "Membre :" $GrpMembers.Name -BackgroundColor red -ForegroundColor black
                        
                    }
                }
                elseif($CheckObject.ObjectClass -eq "user")
                {
                    $y++
                    $worksheet.Cells.Item($y+1,$x+1) = $CheckObject.Name
                    write-host "Utilisateur : " $CheckObject.Name -BackgroundColor red -ForegroundColor Black
                }
            $y++
    	    }
$y++
}

Stop-Transcript

$WorkBook.Save()
$WorkBook.close($true)
