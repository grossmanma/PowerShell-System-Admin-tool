#========================================================================
# created  On: 2/23/2015
# created By: Matthew Grossman 
# Update UPMrest and Prn-restert 9/30/2015
# Update 11/3/2015
# Updated 7/22/2016
# Updated 8/2/2016 adding desktop sync 
# Updated 1/12/2017 adding User office and site. plus removing old button 
# updated 7/11/2017 adding secure print reset button
# adding  10/20/2017 Chrome Update service fix button

#========================================================================  
#----------------------------------------------  
#region Application Functions 
#----------------------------------------------  
  

function OnApplicationLoad { 
	$XMLFile = "ODHD.Options.xml"
	$Script:ParentFolder = Split-Path (Get-Variable MyInvocation -scope 1 -ValueOnly).MyCommand.Definition 
	$XMLFile = Join-Path $ParentFolder $XMLFile
	[XML]$Script:XML = Get-Content $XMLFile

	if($XML.Options.Elevate.Enabled -eq $true){Start-Process -FilePath PowerShell.exe -ArgumentList $MyInvocation.MyCommand.Definition -Verb RunAs} 
$now=Get-Date -format "dd-MMM-yyyy HH:mm"
 Write-output  "$env:UserName ,  $now , Start"| Out-File  -Append   .\logs\access-logs.csv
	
	return $true #return true for success or false for failure
}

function OnApplicationExit {
	
	$script:ExitCode = 0 #Set the exit code for the Packager 
}  

#endregion Application Functions 

#---------------------------------------------- 
# Generated Form Function
#----------------------------------------------
function Call-odhd_pff {

	#----------------------------------------------
	#region Import the Assemblies 
	#----------------------------------------------
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formMain = New-Object System.Windows.Forms.Form
	$groupTools = New-Object System.Windows.Forms.GroupBox 
    $userSelectunlock = New-Object System.Windows.Forms.ToolStripMenuItem 
    $btnsync = New-Object System.Windows.Forms.Button
    $btnNS = New-Object System.Windows.Forms.Button 
    $btnTrust = New-Object System.Windows.Forms.Button 
    $btnWMIR = New-Object System.Windows.Forms.Button 
    $btnSCCMT = New-Object System.Windows.Forms.Button
    $btnprofileUPM1 = New-Object System.Windows.Forms.Button
    $btnUPMRestore = New-Object System.Windows.Forms.Button 
	$btnRestart = New-Object System.Windows.Forms.Button
	$btnGPO = New-Object System.Windows.Forms.Button
    $btnChrome = New-Object System.Windows.Forms.Button
	$btnRA = New-Object System.Windows.Forms.Button 
	$btnRDP = New-Object System.Windows.Forms.Button
    $btnPWSync = New-Object System.Windows.Forms.Button
    $btnSecPRT = New-Object System.Windows.Forms.Button
	$groupInfo = New-Object System.Windows.Forms.GroupBox
	$btnDSA = New-Object System.Windows.Forms.Button
	$btnProcesses = New-Object System.Windows.Forms.Button
	$btnAppv = New-Object System.Windows.Forms.Button 
	$btnEventVwr = New-Object System.Windows.Forms.Button 
	$btnViewUser = New-Object System.Windows.Forms.Button
	$btnServices = New-Object System.Windows.Forms.Button
	$btnDash = New-Object System.Windows.Forms.Button
    $btnHung = New-Object System.Windows.Forms.Button
	$btnSystemInfo = New-Object System.Windows.Forms.Button
	$lvMain = New-Object System.Windows.Forms.ListView
	$btnSearch = New-Object System.Windows.Forms.Button
	$btnSearchuser = New-Object System.Windows.Forms.Button
    $btnPRT = New-Object System.Windows.Forms.Button
	$txtComputer = New-Object System.Windows.Forms.TextBox
	$txtUser = New-Object System.Windows.Forms.TextBox
	$SB = New-Object System.Windows.Forms.StatusBar
	$menu = New-Object System.Windows.Forms.MenuStrip
	$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuView = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuViewUPMT = New-Object System.Windows.Forms.ToolStripMenuItem
	$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip 
	$cmsProcEnd = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsStartupRemove = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAdminAdd = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAdminRemove = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsAppUninstall = New-Object System.Windows.Forms.ToolStripMenuItem
    $userSelectGPM = New-Object System.Windows.Forms.ToolStripMenuItem
    $userSelectunfind = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsSelect = New-Object System.Windows.Forms.ToolStripMenuItem
	$userSelect = New-Object System.Windows.Forms.ToolStripMenuItem
	$GPSelectadd  = New-Object System.Windows.Forms.ToolStripMenuItem
	$GPSelectremove = New-Object System.Windows.Forms.ToolStripMenuItem
	$cmsSelogoffuser = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
	$SBPStatus = New-Object System.Windows.Forms.StatusBarPanel
	$SBPBlog = New-Object System.Windows.Forms.StatusBarPanel
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#---------------------------------------------- 
	
		
	$formMain_Load={
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null 
		$VBMsg = New-Object -COMObject WScript.Shell
        
		if($XML.Options.Domain.Enabled -eq $true){$Domain = $XML.Options.Domain.Default}
		elseif("\\$env:computername" -eq $env:logonserver){$Domain = "Local"}
        else{$Domain = ([DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name}
		Set-FormTitle
	}
       $btnChrome_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*)
                       Get-ComputerName
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
         $a = new-object -comobject wscript.shell
         $ServiceName = "gupdate", "gupdatem"
         $processName = "Chrome", "GoogleUpdate"
                 $a.popup( "Working on $ComputerName" )           
 if(!(Test-Connection -ComputerName $ComputerName -Count 1 -quiet)) {            
  $a.popup( "$ComputerName : Offline" )           
           
 } 
  


 
foreach ($process in $ProcessName) {
    (Get-WmiObject Win32_Process -ComputerName $ComputerName| ?{     $_.ProcessName -match "$process" }).Terminate()



 
if($returnval.returnvalue -eq 0) {
  write-host "The process $ProcessName `($processid`) terminated successfully"
}
else {
  write-host "The process $ProcessName `($processid`) termination has some problems"
}}
 
   

                
             
 foreach($service in $ServiceName) {            
  try {            
   $ServiceObject = Get-WMIObject -Class Win32_Service -ComputerName $ComputerName -Filter "Name='$service'" -EA Stop            
   if(!$ServiceObject) {            
    $a.popup( "$ComputerName : No service found with the name $service")            
               
   }            
   if($ServiceObject.StartMode -eq "Disabled") {            
    $a.popup( "$ComputerName : Service with the name $service already in disabled state"  )          
    Continue            
   }            
               
   Set-Service -ComputerName $ComputerName -Name $service -EA Stop -StartMode Disabled            
   $a.popup( "$ComputerName : Successfully disabled the service $service. Trying to stop it" )           
   if($ServiceObject.Status -eq "Running") {            
    $a.popup( "$ComputerName : $service already in stopped state" )           
    Continue            
   }            
   $retval = $ServiceObject.StopService()            
            
   if($retval.ReturnValue -ne 0) {            
   $a.popup( "$ComputerName : Failed to stop service. Return value is $($retval.ReturnValue)" )           
    Continue            
   }            
               
   $a.popup( "$ComputerName : Stopped service successfully" )  
  if (Test-Path -Path "\\$computername\c$\Program Files (x86)\Google\Update"){
   Remove-Item -Path "\\$computername\c$\Program Files (x86)\Google\Update" -Recurse -Force -Confirm:$False
   }
   Else { Remove-Item -Path "\\$computername\c$\Program Files\Google\Update" -Recurse -Force -Confirm:$False
    $a.popup("Chrome Update Fix was Successful."  )    
   }
               
  } catch {            
   $a.popup("$ComputerName : Failed to query $service. Details : $_"  )          
   Continue            
  }            
             
 }  } 

       $btnNS_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*)
        Start-Process C:\windows\System32\wscript.exe '.\WSMaker\invisible.vbs'


       }

       $btnSecPrt_click={
remove-ContextMenu(get-Variable Userselect*) 
remove-ContextMenu(get-Variable GPSelect*)
Get-ComputerName
#Load VB module
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$a = new-object -comobject wscript.shell
if(-not(test-connection -ComputerName $computername -Quiet)){
$a.popup("$computername  is not online please check the computer name and try again")
return
}

 $service = Invoke-Command -ComputerName $computername {powershell.exe Get-Service  -name "LRSDRVX"} 
 if ($service) {
  $VerifyServiceStopped1 = Invoke-Command -ComputerName $computername {powershell.exe Get-Service  -name "LRSDRVX"| Where-Object {$_.status -eq "Stopped"} | select -last 1}
    if ($VerifyServiceStopped1) {
         $a.popup("Secure Print is install checking that the service is running on $computername ")
         $a.popup("Secure Print service is running on $computername  Have the User run the c:\SecPrt.bat ")

       Invoke-Command   -ComputerName $computername   { Function Send-TSMessageBox
{
    param([string]$Title = "Title", [string]$Message = "Message", [int]$ButtonSet = 0, [int]$Timeout = 0, [bool]$WaitResponse = $false)
    
        $Signature = @"
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSSendMessage(
            IntPtr hServer,
            [MarshalAs(UnmanagedType.I4)] int SessionId,
            String pTitle,
            [MarshalAs(UnmanagedType.U4)] int TitleLength,
            String pMessage,
            [MarshalAs(UnmanagedType.U4)] int MessageLength,
            [MarshalAs(UnmanagedType.U4)] int Style,
            [MarshalAs(UnmanagedType.U4)] int Timeout,
            [MarshalAs(UnmanagedType.U4)] out int pResponse,
            bool bWait);
            
            [DllImport("kernel32.dll")]
            public static extern uint WTSGetActiveConsoleSessionId();
"@

            [int]$TitleLength = $Title.Length;
            [int]$MessageLength = $Message.Length;
            [int]$Response = 0;
                                    
            $MessageBox = Add-Type -memberDefinition $Signature -name "WTSAPISendMessage" -namespace "WTSAPI" -passThru   
            $SessionId = $MessageBox::WTSGetActiveConsoleSessionId()
            
            $MessageBox::WTSSendMessage(0, $SessionId, $Title, $TitleLength, $Message, $MessageLength, $ButtonSet, $Timeout, [ref] $Response, $WaitResponse)

             "C:\Program Files\LRS\VPSX Printer Driver Management\vspa.exe connect" >> c:\SecPrt.bat
            
            $Response 

            
            
}
Send-TSMessageBox -Title "Secure Print"  -Message "System requires run the connect command? Please run c:\SecPrt.bat " -ButtonSet 0 -Timeout 60} 
         a.popup("please check for printer now  $computername"  )
         return
    } 
 else{
  $a.popup("Secure Printis service is not running on $computername  starting Service")
     try
     {


        Invoke-Command -ComputerName $computername {powershell.exe Start-Service  -name "LRSDRVX"} 1> $null
        if(!$?){ $a.popup("Was unable to start the service on $computername please check $computername")
        return
}
        $a.popup("Verifying that Secure Print  service is started $computername")
        $VerifyServiceStopped1 = Invoke-Command -ComputerName $computername {powershell.exe Get-Service  -name "LRSDRVX"| Where-Object {$_.status -eq "Stopped"} | select -last 1}
          if ($VerifyServiceStopped1) {
         Invoke-Command   -ComputerName $computername   { Function Send-TSMessageBox
{
    param([string]$Title = "Title", [string]$Message = "Message", [int]$ButtonSet = 0, [int]$Timeout = 0, [bool]$WaitResponse = $false)
    
        $Signature = @"
        [DllImport("wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSSendMessage(
            IntPtr hServer,
            [MarshalAs(UnmanagedType.I4)] int SessionId,
            String pTitle,
            [MarshalAs(UnmanagedType.U4)] int TitleLength,
            String pMessage,
            [MarshalAs(UnmanagedType.U4)] int MessageLength,
            [MarshalAs(UnmanagedType.U4)] int Style,
            [MarshalAs(UnmanagedType.U4)] int Timeout,
            [MarshalAs(UnmanagedType.U4)] out int pResponse,
            bool bWait);
            
            [DllImport("kernel32.dll")]
            public static extern uint WTSGetActiveConsoleSessionId();
"@

            [int]$TitleLength = $Title.Length;
            [int]$MessageLength = $Message.Length;
            [int]$Response = 0;
                                    
            $MessageBox = Add-Type -memberDefinition $Signature -name "WTSAPISendMessage" -namespace "WTSAPI" -passThru   
            $SessionId = $MessageBox::WTSGetActiveConsoleSessionId()
            
            $MessageBox::WTSSendMessage(0, $SessionId, $Title, $TitleLength, $Message, $MessageLength, $ButtonSet, $Timeout, [ref] $Response, $WaitResponse)

             "C:\Program Files\LRS\VPSX Printer Driver Management\vspa.exe connect" >> c:\SecPrt.bat
            
            $Response 

            
            
}
Send-TSMessageBox -Title "Secure Print"  -Message "System requires run the connect command? Please run c:\SecPrt.bat " -ButtonSet 0 -Timeout 60} 
         a.popup("please check for printer now  $computername"  )
         return
    } 
}
catch  
{
}}


       
      
   
 } 
 Else {

  $a.popup("Secure Print is not install on $computername. Installing now. this can take couple of minutes ")

   $pro = get-wmiobject win32_computersystem -computer $computername | select-object systemtype
    if($pro -eq "x86*"){
    $computername  | out-file \\dc1-p-a-scsmss1\portal-tools$\SecPRT\x86\$computername} 
    else 
    { $computername  | out-file \\dc1-p-a-scsmss1\portal-tools$\SecPRT\x64\$computername} 
 
 sleep 45
 $a.popup("checking the install $computername. please wait ")
  Invoke-Command -ComputerName  $ComputerName {Stop-Process -processname "CcmExec" -force} 
  Invoke-Command -ComputerName  $ComputerName {start-Service -Name CcmExec} 
 sleep 45
          $a.popup( "starting Secure Print  $computername.")

Invoke-Command -ComputerName $computername {powershell.exe Start-Service  -name "LRSDRVX"}
sleep 90
} 

  $service = Invoke-Command -ComputerName $computername {powershell.exe Get-Service  -name "LRSDRVX"}
 $a.popup("Secure Print is install checking that the service is running on $computername ")
 
 if ($service) {
                
     $a.popup("Secure Print is install $computername please restart $computername")
     }
}


       $btnHung_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*)
                       Get-ComputerName
            
            $computername  | out-file .\Hung-Session\$computername

}


		$btnsync_Click={

         		   #Load VB module
                   [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                   $a = new-object -comobject wscript.shell
                    Write-Host $txtUser.text
         if($txtUser.text -eq ""){
                       $a.popup( " User textbox cannot empty ")
              }else{
                     remove-ContextMenu(get-Variable Userselect*) 
                     remove-ContextMenu(get-Variable GPSelect*)
                     Get-UserName
                   $SBPStatus.Text = "Getting ready to sync desktop for $username..."
            [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
         $a = new-object -comobject wscript.shell
         $intAnswer = $a.popup( " Is this a DC1 user $username Please comfirm in Citrix Director", `
0,"Rename Profile",4) 
If ($intAnswer -eq 6) { 
if(!(Test-Path  \\DFS-server\Redir\$username )){
  

      $a.popup("Desktop not found for $username ")
}
else{
        $SBPStatus.Text = "starting desktop sync for $username at DC1..."
		$a.popup( " Confirm DC1 user $username ")
		$a.popup( " Please sure make all unwanted items are remobe from $username desktop. Before clicking OK")
        robocopy \\dfs-server1\Redir\$username \\dfs-server\Redir\$username /MIR
                $SBPStatus.Text = "Completed desktop sync for $username at DC1..."
        $a.popup( " Desktop is now in-sync for $username ")

    }}
else {
if(!(Test-Path  \\dfs-server\Redir\$username ))
  {

      $a.popup("Desktop not found for $username ")
}
else{
    $SBPStatus.Text = "starting desktop sync for $username at DC2..."
    $a.popup( " Confirm DC2 user $username ")
    $a.popup( " Please make sure all unwanted items are remobe from $username desktop. Before clicking OK")
    robocopy \\dfs-server\Redir\$username \\dfs-server\Redir\$username /MIR
    $SBPStatus.Text = "Completed desktop sync for $username at DC1..."
    $a.popup( " Desktop is now in-sync for $username ")
    
   } } }}
	
	$btnGPO_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
                $SBPStatus.Text = "Group Policy..."
		Import-Module -Name grouppolicy
		Get-Command -Module grouppolicy
               if (!$username){ 
               Get-UserName
               Get-ComputerName

                    Get-GPResultantSetOfPolicy -user "corp\$username" -Computer "$computername" -ReportType Html -Path "C:\Windows\Temp\$username.html"
                    Start-Process  -filepath "C:\Program Files (x86)\Internet Explorer\IEXPLORE.EXE"  "C:\Windows\Temp\$username.html"
    }
    else
    {             $txtComputer.Text = $lvMain.SelectedItems[0].Text
                        Get-ComputerName  
                    Get-GPResultantSetOfPolicy  -Computer "$computername" -ReportType Html -Path "C:\Windows\Temp\$computername.html"
                    Start-Process  -filepath "C:\Program Files (x86)\Internet Explorer\IEXPLORE.EXE"  "C:\Windows\Temp\$computername.html"

}}

	
	
	$btnSearch_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		Initialize-Listview
		$SBPStatus.Text = "Retrieving Computers..."
		Update-ContextMenu (Get-Variable cmsSelect*)
		$Properties = $XML.Options.Search.Property 
		$Properties | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
		$Info = Get-RPADComputer
		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.Properties.(($Col0).ToLower()))
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				[String]$SubItem = $_.Properties.(($Field).ToLower())
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
		$SBPStatus.Text = "Ready"
	}

	$cmsSelect_Click={
		if ($lvMain.SelectedItems.Count -gt 1){$vbError = $vbmsg.popup("You may only select one computer at a time.",0,"Error",0)}
		else{$txtComputer.Text = $lvMain.SelectedItems[0].Text}
	}
	
	       $btnSearchuser_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
              Get-UserName
              Initialize-Listview
              $SBPStatus.Text = "Retrieving User...."
              Update-ContextMenu (Get-Variable userSelect*)
              $Properties = $XML.Options.SearchU.Property | %{Add-Column $_}
              Resize-Columns
              $Col0 = $lvMain.Columns[0].Text
        $Info = @()
        $Name = $username + "*"
        $Users = Get-ADUser -Filter 'GivenName  -like $name -or  SamAccountName  -like $name -or Surname -like $name'  -Properties * | Select SamAccountName, DisplayName, Enabled, LockedOut, LastLogonDate, office, site  | sort name 
        
        Foreach ($User in $Users)
            {
            $Office = $User.office 
            $Group = (Get-ADGroup LOC_SC_$Office).distinguishedname
            $DC1 = (Get-ADGroup -Identity "LOC_DC_DC01" -Properties members).members
            $DC2 = (Get-ADGroup -Identity "LOC_DC_DC02" -Properties members).members
            If ($DC1 -contains $Group)
                {$site = "DC1"}
            If ($DC2 -contains $Group)
                {$site = "DC2"}
            
            $Info += [pscustomobject]@{
                SamAccountName = $User.SamAccountName
                DisplayName = $user.displayname
                Enabled = $User.enabled
                LockedOut = $User.LockedOut
                LastLogonDate = $User.lastlogonDate
                office = $User.office
                site = $Site
                }

            }
    
              $Info | %{
                     $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
                     ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
                           $Field = $Col.Text
                           $SubItem = $_.$Field | out-string
                           if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
                           else{$Item.SubItems.Add("")}
                     }
                     $lvMain.Items.Add($Item)
              }
              $SBPStatus.Text = "Ready"
       } 

	$userSelect_Click={
		if ($lvMain.SelectedItems.Count -gt 1){$vbError = $vbmsg.popup("You may only select one User at a time.",0,"Error",0)}
		else{$txtuser.Text = $lvMain.SelectedItems[0].Text
        } &$btnSearchuser_Click

        

	}       

	$cmsSelect_Click={
		if ($lvMain.SelectedItems.Count -gt 1){$vbError = $vbmsg.popup("You may only select one computer at a time.",0,"Error",0)}
		else{ 
                $txtComputer.Text = $lvMain.SelectedItems[0].text
                Get-ComputerName} &$btnSearch_Click
	} 

	$userSelectunlock_Click={
         		if ($lvMain.SelectedItems.Count -gt 1){$vbError = $vbmsg.popup("You may only select one User at a time.",0,"Error",0)}
		else{$txtuser.Text = $lvMain.SelectedItems[0].Text}
                #Load VB module
                   [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                   $a = new-object -comobject wscript.shell
             $SBPStatus.Text = "unlocking  $username Account..."
             Unlock-ADAccount -Identity $username
             $a.popup("The $username has been unlocked")        
   }

  $userSelectunfind_Click={     
           		if ($lvMain.SelectedItems.Count -gt 1){$vbError = $vbmsg.popup("You may only select one User at a time.",0,"Error",0)}
		else{$txtuser.Text = $lvMain.SelectedItems[0].Text}   
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-UserName
		Initialize-Listview
		$SBPStatus.Text = "Retrieving User Info...."
		Update-ContextMenu (Get-Variable userSelect*)
		$Properties = $XML.Options.SearchF.Property | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
		$Info = .\Get-LockedOutUser.ps1 -UserName $username -StartTime (Get-Date).AddDays(-1) | Select-Object ClientName, TimeCreated, UserName   | sort name 

		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				$SubItem = $_.$Field | out-string
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
		$SBPStatus.Text = "Ready"
	}

  
  
  

   	$userSelectGPM_Click={
         
         		if ($lvMain.SelectedItems.Count -gt 1){$vbError = $vbmsg.popup("You may only select one User at a time.",0,"Error",0)}
		else{$txtuser.Text = $lvMain.SelectedItems[0].Text}
        remove-ContextMenu(get-Variable Userselect*) 
         Get-UserName
         Initialize-Listview
         $SBPStatus.Text = "Looking up Group Membership for $username..." 
         Update-ContextMenu (Get-Variable GPSelect*)                        	
		$Properties = $XML.Options.SearchGP.Property | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
		$Info = Get-ADPrincipalGroupMembership  $username  | Select name | sort name 

		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				$SubItem = $_.$Field | out-string
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
         
		$SBPStatus.Text = "Ready"
	}
      
     $GPSelectadd_click ={
     [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
     [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
     Function Add-Object([string]$obj, [string]$group){
			$dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
			$root = $dom.GetDirectoryEntry()
			$search = [System.DirectoryServices.DirectorySearcher]$root
			$search.Filter = "(CN=$obj)"
			$result = $search.FindOne()
			$user = [ADSI]$result.path
			$user = $user.distinguishedName
			$search.Filter = "(CN=$group)"
			$result = $search.FindOne()
			$groupToADD = [ADSI]$result.path
			Try{
			$groupToADD.member.add("$user")
			$groupToADD.setinfo()
			[Windows.Forms.MessageBox]::Show("Object successfully added to $group.")}
			Catch{[Windows.Forms.MessageBox]::Show("Permission is denied.")}

}


function Return-DropDown{
		
            $Form = New-Object System.Windows.Forms.Form
			$Form.width = 300
			$Form.height = 250
			$Form.Text = "Add User to AD Group"
			$Form.maximumsize = New-Object System.Drawing.Size(300,250)
			$Form.startposition = "centerscreen"
			$Form.KeyPreview = $True
			$Form.Add_KeyDown({if ($_.KeyCode -eq "Enter"){Add-Object $textboxOBJ.Text}})
			$Form.Controls.Add($DropDown)
			$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Form.Close()}})


			
			$OKButton = new-object System.Windows.Forms.Button
			$OKButton.Location = new-object System.Drawing.Size(60,150)
			$OKButton.Size = new-object System.Drawing.Size(80,20)
			$OKButton.Text = "OK"
			$OKButton.Add_Click({Add-Object $textboxOBJ.Text $textboxGROUP.Text})
			
			$CancelButton = New-Object System.Windows.Forms.Button
			$CancelButton.Location = New-Object System.Drawing.Size(150,150)
			$CancelButton.Size = New-Object System.Drawing.Size(80,20)
			$CancelButton.Text = "Cancel"
			$CancelButton.Add_Click({$Form.Close()})
			
			
			$textboxOBJ = new-object System.Windows.Forms.TextBox
			$textboxOBJ.Location = new-object System.Drawing.Size(120,10)
			$textboxOBJ.Size = new-object System.Drawing.Size(150,20)
            $textboxOBJ.text = $username

			$labelOBJ = new-object System.Windows.Forms.Label
			$labelOBJ.Location = new-object System.Drawing.Size(10,10)
			$labelOBJ.size = new-object System.Drawing.Size(50,20)
			$labelOBJ.Text = "User"
			
		  
            $textboxGROUP = new-object System.Windows.Forms.ComboBox
			$textboxGROUP.Location = new-object System.Drawing.Size(90,50)
			$textboxGROUP.Size = new-object System.Drawing.Size(160,20)  
            $textboxgroup.Items.Add($textboxGROUP.Text ) 


			$labelGROUP = new-object System.Windows.Forms.Label
			$labelGROUP.Location = new-object System.Drawing.Size(10,50)
			$labelGROUP.size = new-object System.Drawing.Size(70,20)
			$labelGROUP.Text = "Group"

			[array]$DropDownArray = (Get-ADGroup -filter {GroupCategory -eq "Security" -and GroupScope -eq "Global"} | Select name | sort name )
			
			ForEach ($Item in $DropDownArray  -replace '[@{name= }]','') {
				$textboxGROUP.Items.Add($Item)  |  Out-Null}


			
			$Form.Controls.Add($CancelButton)
			$Form.Controls.Add($labelGROUP)
			$Form.Controls.Add($textboxGROUP)
			$Form.Controls.Add($labelOBJ)
			$Form.Controls.Add($textboxOBJ)
			$Form.Controls.Add($OKButton)
			$Form.Add_Shown({$Form.Activate()})
			$Form.ShowDialog()
}     

Return-DropDown

   }
	
    $GPSelectremove_click ={ 
 
  $SBPStatus.Text = "removing user from  Group ..."
  $GPname = $lvMain.SelectedItems.text 
  #remove-adgroupmember -Identity $GPname -Members $UserName
  $ScriptBlockContent = { param ($gpname)  remove-adgroupmember -Identity $GPname  -Members $UserName}
  $ScriptBlockContent
  Remove-SelectedItems
   $SBPStatus.Text = "User has been removed from  Group ..."
 } 	

	$btnSystemInfo_Click={ 
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		Initialize-Listview
		$SBPStatus.Text = "Retrieving System Information..."
		Update-ContextMenu (Get-Variable cmsSystem*)
		
		$SysError = $null
		$sysComp = Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName -ErrorVariable SysError
		Start-Sleep -m 250
		if($SysError){$SBPStatus.Text = "[$ComputerName] $SysError"}
		else{
			$sysComp2 = Get-WmiObject Win32_ComputerSystemProduct -ComputerName $ComputerName
			$sysOS = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName
			$sysBIOS = Get-WmiObject Win32_BIOS -ComputerName $ComputerName
			$sysCPU = Get-WmiObject Win32_Processor -ComputerName $ComputerName
			$sysRAM = Get-WmiObject Win32_PhysicalMemory -ComputerName $ComputerName
			$sysNAC = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "IPEnabled='True'"
			$sysMon = Get-WmiObject Win32_DesktopMonitor -ComputerName $ComputerName
			$sysVid = Get-WmiObject Win32_VideoController -ComputerName $ComputerName
			$sysOD = Get-WmiObject Win32_CDROMDrive -ComputerName $ComputerName
			$sysHD = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName
			$sysProc = Get-WmiObject Win32_Process -ComputerName $ComputerName
		
			if ($XML.Options.SystemInfo.AntiVirus.Enabled -eq $true){
				$sysAV = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct -ComputerName $ComputerName
				switch ($sysAV.ProductState) {
					"262144" {$DefStatus = "Up to date"  ;$RTStatus = "Disabled"}
				    "262160" {$DefStatus = "Out of date" ;$RTStatus = "Disabled"}
				    "266240" {$DefStatus = "Up to date"  ;$RTStatus = "Enabled"}
				    "266256" {$DefStatus = "Out of date" ;$RTStatus = "Enabled"}
				    "393216" {$DefStatus = "Up to date"  ;$RTStatus = "Disabled"}
				    "393232" {$DefStatus = "Out of date" ;$RTStatus = "Disabled"}
				    "393488" {$DefStatus = "Out of date" ;$RTStatus = "Disabled"}
				    "397312" {$DefStatus = "Up to date"  ;$RTStatus = "Enabled"}
				    "397328" {$DefStatus = "Out of date" ;$RTStatus = "Enabled"}
				    "397584" {$DefStatus = "Out of date" ;$RTStatus = "Enabled"}
					default  {$DefStatus = "Unknown" ;$RTStatus = "Unknown"}
				} 
			}
	
			if ($XML.Options.SystemInfo.General.DomainLocation.Enabled -eq $true -AND $Domain -ne 'Local'){
				$Script:ComputerName = $ComputerName.Split('.')[0]
				$sysOU = Get-RPADComputer
			}
	
			"Property","Value" | %{Add-Column $_}
			
			if ($XML.Options.SystemInfo.General.Enabled -eq $true){
				$Item = New-Object System.Windows.Forms.ListViewItem("General")
				$Item.BackColor = "Black"
				$Item.ForeColor = "White"
				$lvMain.Items.Add($Item)
				if($XML.Options.SystemInfo.General.ComputerName.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("Computer Name")
					$Item.SubItems.Add($sysComp.Name)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.General.DomainLocation.Enabled -eq $true -AND $Domain -ne 'Local'){
					$OU = $sysOU.Path.Substring($sysOU.Path.IndexOf(',')+1)
					$Item = New-Object System.Windows.Forms.ListViewItem("Computer OU")
					$Item.SubItems.Add($OU)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.General.CurrentUser.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("User")
					if($SysComp.UserName -ne $null){$Item.SubItems.Add($sysComp.UserName)}
					else{$Item.SubItems.Add("")}					
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.General.LogonTime.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("User Logon")
					if($sysProc | ?{$_.Name -eq "explorer.exe"}){
						$UserLogonDT = ($sysProc | ?{$_.Name -eq "explorer.exe"} | Sort CreationDate | Select-Object -First 1).CreationDate
						$UserLogon = [System.Management.ManagementDateTimeconverter]::ToDateTime($UserLogonDT).ToString()
						$Item.SubItems.Add($UserLogon)
					}else{
						$Item.SubItems.Add("N/A")
					}
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.General.ScreenSaverTime.Enabled -eq $true){
					
					$Item = New-Object System.Windows.Forms.ListViewItem("Screensaver Time")
					if($sysProc | ?{$_.Name -match ".scr"}){
						$ScreensaverTime = ($sysProc | ?{$_.Name -match ".scr"} | Sort CreationDate | Select-Object -First 1).CreationDate
						$Screensaver = [System.Management.ManagementDateTimeconverter]::ToDateTime($ScreensaverTime).ToString()
						$Item.SubItems.Add($Screensaver)
					}else{
						$Item.SubItems.Add("N/A")
					}
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.General.LastRestart.Enabled -eq $true){
					$LastBootUpTime = [System.Management.ManagementDateTimeconverter]::ToDateTime($sysOS.LastBootUpTime).ToString()
					$Item = New-Object System.Windows.Forms.ListViewItem("Last Restart")
					$Item.SubItems.Add($LastBootUpTime)
					$lvMain.Items.Add($Item)
				}
			}
			
			if ($XML.Options.SystemInfo.Build.Enabled -eq $true){
				$Item = New-Object System.Windows.Forms.ListViewItem("Build")
				$Item.BackColor = "Black"
				$Item.ForeColor = "White"
				$lvMain.Items.Add($Item)
				if($XML.Options.SystemInfo.Build.Manufacturer.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("Manufacturer")
					$Item.SubItems.Add($sysComp.Manufacturer)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.Build.Model.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("Model")
					$Item.SubItems.Add($sysComp.Model)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.Build.Chassis.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("Chassis")
					$Item.SubItems.Add($sysComp2.Version)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.Build.Serial.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("Serial")
					$Item.SubItems.Add($sysBIOS.SerialNumber)
					$lvMain.Items.Add($Item)
				}
			}
			
			if ($XML.Options.SystemInfo.Hardware.Enabled -eq $true){
				$Item = New-Object System.Windows.Forms.ListViewItem("Hardware")
				$Item.BackColor = "Black"
				$Item.ForeColor = "White"
				$lvMain.Items.Add($Item)
				if($XML.Options.SystemInfo.Hardware.CPU.Enabled -eq $true){
					$sysCPU | %{
					$Item = New-Object System.Windows.Forms.ListViewItem("CPU")
					$Item.SubItems.Add($sysCPU.Name.Trim())
					$lvMain.Items.Add($Item)
					}
				}
				if($XML.Options.SystemInfo.Hardware.RAM.Enabled -eq $true){
					$tRAM = "{0:N2} GB Usable - " -f $($sysComp.TotalPhysicalMemory / 1GB)
					$sysRAM | %{$tRAM += "[$($_.Capacity / 1GB)] "}
					$Item = New-Object System.Windows.Forms.ListViewItem("RAM")
					$Item.SubItems.Add($tRAM)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.Hardware.HD.Enabled -eq $true){
					$sysHD | ?{$_.DriveType -eq 3} | %{
						$HDInfo = "{0:N1} GB Free / {1:N1} GB Total" -f ($_.FreeSpace / 1GB), ($_.Size / 1GB)
						$Item = New-Object System.Windows.Forms.ListViewItem("HD")
						$Item.SubItems.Add($HDinfo)
						$lvMain.Items.Add($Item)
					}
				}
				if($XML.Options.SystemInfo.Hardware.OpticalDrive.Enabled -eq $true){
					$sysOD | %{
					$Item = New-Object System.Windows.Forms.ListViewItem("Optical Drive")
					$Item.SubItems.Add("[$($sysOD.Drive)] $($sysOD.Caption)")
					$lvMain.Items.Add($Item)
					}
				}
				if($XML.Options.SystemInfo.Hardware.GPU.Enabled -eq $true){
					$sysVid | ?{$_.AdapterRAM -gt 0} | %{
					$Item = New-Object System.Windows.Forms.ListViewItem("GPU")
					$Item.SubItems.Add($_.Name)
					$lvMain.Items.Add($Item)
					}
				}
				if($XML.Options.SystemInfo.Hardware.Monitor.Enabled -eq $true){
					$Monitors = $null
					$sysMON | %{$Monitors += "[{0} x {1}] " -f $_.ScreenWidth,$_.ScreenHeight}
					$Item = New-Object System.Windows.Forms.ListViewItem("Monitor(s)")
					$Item.SubItems.Add($Monitors)
					$lvMain.Items.Add($Item)
				}
			}
			
			if ($XML.Options.SystemInfo.OS.Enabled -eq $true){
				$Item = New-Object System.Windows.Forms.ListViewItem("Operating System")
				$Item.BackColor = "Black"
				$Item.ForeColor = "White"
				$lvMain.Items.Add($Item)
				if($XML.Options.SystemInfo.OS.OS.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("OS Name")
					$Item.SubItems.Add($sysOS.Caption)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.OS.ServicePack.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("Service Pack")
					$Item.SubItems.Add($sysOS.CSDVersion)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.OS.Architecture.Enabled -eq $true){
					$Item = New-Object System.Windows.Forms.ListViewItem("OS Architecture")
					$Item.SubItems.Add($sysComp.SystemType)
					$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.OS.ImageDate.Enabled -eq $true){
					$InstallDate = [System.Management.ManagementDateTimeconverter]::ToDateTime($sysOS.InstallDate).ToString()
					$Item = New-Object System.Windows.Forms.ListViewItem("Install Date")
					$Item.SubItems.Add($InstallDate)
					$lvMain.Items.Add($Item)
				}
			}
			
			if ($XML.Options.SystemInfo.IPConfig.Enabled -eq $true){
				$Item = New-Object System.Windows.Forms.ListViewItem("Network Adapters")
				$Item.BackColor = "Black"
				$Item.ForeColor = "White"
				$lvMain.Items.Add($Item)
				$sysNAC | %{
					if($XML.Options.SystemInfo.IPConfig.Description.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("Description")
						$Item.SubItems.Add($_.Description)
						$lvMain.Items.Add($Item)
					}
					if($XML.Options.SystemInfo.IPConfig.IPAddress.Enabled -eq $true){
						$IPinfo = $null
						ForEach ($IP in $_.IPAddress){$IPinfo += "$IP "}
						$Item = New-Object System.Windows.Forms.ListViewItem("IP Address")
						$Item.SubItems.Add($IPinfo)
						$lvMain.Items.Add($Item)
					}
					if($XML.Options.SystemInfo.IPConfig.MACAddress.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("MAC Address")
						$Item.SubItems.Add($_.MACAddress)
						$lvMain.Items.Add($Item)
					}
					if($XML.Options.SystemInfo.IPConfig.DHCPEnabled.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("DHCP Enabled")
						$Item.SubItems.Add($_.DHCPEnabled.ToString())
						$lvMain.Items.Add($Item)
					}
					if($XML.Options.SystemInfo.IPConfig.DHCPServer.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("DHCP Server")
						$Item.SubItems.Add($_.DHCPServer)
						$lvMain.Items.Add($Item)
					}
					if($XML.Options.SystemInfo.IPConfig.DNSDomain.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("DNS Domain")
						$Item.SubItems.Add($_.DNSDomain)
						$lvMain.Items.Add($Item)
					}
				}
			}
			if ($XML.Options.SystemInfo.Antivirus.Enabled -eq $true){
				$Item = New-Object System.Windows.Forms.ListViewItem("AntiVirus")
				$Item.BackColor = "Black"
				$Item.ForeColor = "White"
				$lvMain.Items.Add($Item)
				if($XML.Options.SystemInfo.Antivirus.Name.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("AV Name")
						$Item.SubItems.Add($sysAV.DisplayName)
						$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.Antivirus.DefinitionStatus.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("Definition Status")
						$Item.SubItems.Add($DefStatus)
						$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.Antivirus.RealTimeProtection.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("Real-Time Protection")
						$Item.SubItems.Add($RTStatus)
						$lvMain.Items.Add($Item)
				}
				if($XML.Options.SystemInfo.Antivirus.Executable.Enabled -eq $true){
						$Item = New-Object System.Windows.Forms.ListViewItem("Executable")
						$Item.SubItems.Add($sysAV.PathToSignedProductExe)
						$lvMain.Items.Add($Item)
				}
			}
			
			$lvMain.Columns[0].Width = "120"
			$lvMain.Columns[1].Width = ($lvMain.Width - ($lvMain.Columns[0].Width + 22))
			$SBPStatus.Text = "Ready"
		}
	}
	
	$menuFileExit_Click={
		$formMain.Close()
	}
	
	$btnProcesses_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		Initialize-Listview
		$SBPStatus.Text = "Retrieving Processes..."
		Update-ContextMenu (Get-Variable cmsProc*)
		$XML.Options.Processes.Property | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
		$Info = Get-WmiObject win32_process -ComputerName $ComputerName -ErrorVariable SysError | Select-Object ProcessID,Name,Executablepath,@{n='Owner';e={$_.GetOwner().User}} | sort name
		Start-Sleep -m 250
		if($SysError){$SBPStatus.Text = "[$ComputerName] $SysError"}
		else{
		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				$SubItem = $_.$Field
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
		$SBPStatus.Text = "Ready"
		}
	}
	
	$cmsProcEnd_Click={$Item = 

		foreach ($Sel in $lvMain.SelectedItems){($Info | ?{$_.ProcessID -eq $Sel.Text})}
        (Get-WmiObject Win32_Process -ComputerName $ComputerName | ?{$_.ProcessID -eq $Sel.Text}).Terminate()
	Remove-SelectedItems
	}
	
	$btnappv_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		Initialize-Listview
		$SBPStatus.Text = "Retrieving App-v Packages..."
		Update-ContextMenu (Get-Variable cmsApp*)
		$XML.Options.APPvItems.Property | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        Invoke-Command -ComputerName $ComputerName {Set-ExecutionPolicy Unrestricted}
		$Info =  Invoke-Command -ComputerName $ComputerName {Get-AppvClientPackage -all}| Select-Object Name, Version,PercentLoaded, Path | sort name
		Start-Sleep -m 250
		if($SysError){$SBPStatus.Text = "[$ComputerName] $SysError"}
		else{
		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				$SubItem = $_.$Field
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
		$SBPStatus.Text = "Ready"
		}
	}
	
    $cmsAppUninstall_Click={
#Load VB module
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$a = new-object -comobject wscript.shell
		
 $name = $lvMain.SelectedItems.text 

 $ScriptBlockContent = { param ($name) Remove-AppvClientPackage  -name $name}

    Invoke-Command -ComputerName  $ComputerName -ScriptBlock $ScriptBlockContent -ArgumentList $name
    Invoke-Command -ComputerName  $ComputerName {Stop-Process -processname "CcmExec" -force} 
    Invoke-Command -ComputerName  $ComputerName {start-Service -Name CcmExec} 
    sleep -Seconds 10
Invoke-WMIMethod -ComputerName $ComputerName -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}"
	
$a.popup("The App-v package $name has been remove from $ComputerName SCCM should redeploy the package in the next couple of minutes")
Remove-SelectedItems
}       	
	   




        $btnSCCMT_click={  
            remove-ContextMenu(get-Variable Userselect*) 
            remove-ContextMenu(get-Variable GPSelect*)               
               Get-ComputerName
		       Initialize-Listview
		       $SBPStatus.Text = "Starting the SCCM Tool ..."
                Start-Process ".\Cireson\RemoteManage\ConfigMgrClientTools.exe"
    }


   


	$btnEventVwr_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		Initialize-Listview
		$SBPStatus.Text = "Retrieving Event Viewer Items..."
				Get-ComputerName
		EventVwr $ComputerName
	}
	
	$btnViewUser_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		Initialize-Listview
		$SBPStatus.Text = "Retrieving Users and Groups from $computername..."
		LUsrMgr.msc /Computer:$ComputerName
		$SBPStatus.Text = "Ready"
		
	}
	
	$btnDSA_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
        c:\windows\system32\dsa.msc
		}	
	
	$btnRDP_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
#Load VB module
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
# Collect Thin Host Name 
$servers = $ComputerName
$a = new-object -comobject wscript.shell
Foreach($s in $servers)

  {
  #Ping Thin Client
    if(!(Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet))
      # If not able to ping error message
      {$b = $a.popup("Problem exists connecting to $s check the Thin Client Name",4)} 

       Else
    #checking if Remoteregustry service is running
       {
       $arrService = Get-Service -ComputerName $s -Name RemoteRegistry
    # if remoteregistry is not started then started it
       if ($arrService.Status -ne "started"){Get-Service -ComputerName $s -Name RemoteRegistry | Start-Service
              #Message letting you it connecting to remote host
              $b = $a.popup("Please wait while connecting to $s",4)}
       # looking for the reg key
       $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $s) 
       $ref = $regKey.OpenSubKey("SOFTWARE\Microsoft\MSLicensing\Store\LICENSE000"); 
       # if not found stop remoteregistry service and pop-upp messges letting you know it could find the key
       if (!$ref) {Get-Service -ComputerName $s -Name RemoteRegistry | Stop-Service
        $b = $a.popup("*LICENSE* key does not exist $s",4)
       }
       else {
       # look up the key and delete it. stop the remoteregistry service.
        $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $s) 
        $reg = $regKey.OpenSubKey('SOFTWARE\Microsoft\MSLicensing\Store',$true )
        $reg.GetSubKeyNames() | Where { $_ -like "*LICENSE*" } | ForEach {
         "Deleting subkey $_"
         $reg.DeleteSubKeyTree($_)
            Get-Service -ComputerName $s -Name RemoteRegistry | Stop-Service
            # Completion messgae 
             $b = $a.popup("*LICENSE* key deleted $s",4)
      }}

   }} # end if

   }
   
    $btnprt_click ={
            remove-ContextMenu(get-Variable Userselect*) 
            remove-ContextMenu(get-Variable GPSelect*) 
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $a = new-object -comobject wscript.shell
                Get-ComputerName
 
    Invoke-Command -ComputerName  $ComputerName {Stop-Process -processname "spoolsv" -force} 
    Invoke-Command -ComputerName  $ComputerName {Stop-Process -processname "CpSvc" -force}
    Invoke-Command -ComputerName  $ComputerName {Stop-Process -processname "PrintIsolationHost" -force}
    Invoke-Command -ComputerName  $ComputerName {Stop-Process -processname "LMUD1N4Z" -force}


    Invoke-Command -ComputerName  $ComputerName {start-Service -Name Spooler} 
    Invoke-Command -ComputerName  $ComputerName {start-Service -Name CpSvc}


    If ((Get-Process -ComputerName $ComputerName -Name spoolsv).responding -eq 'False')
{
    Invoke-Command -ComputerName  $ComputerName {start-Service -Name Spooler} 
}

    If ((Get-Process -ComputerName $ComputerName -Name CpSvc).responding -eq 'False')
{
    Invoke-Command -ComputerName  $computername {start-Service -Name CpSvc}
}
$a.popup("Print Spooler has been restarted on $ComputerName") }

    $btnprt_click ={
            remove-ContextMenu(get-Variable Userselect*) 
            remove-ContextMenu(get-Variable GPSelect*) 
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $a = new-object -comobject wscript.shell
                Get-ComputerName

 
    Invoke-Command -ComputerName  $ComputerName {Start-Process "C:\Program Files\LRS\VPSX Printer Driver Management\vspa.exe" reconnect} 

    }
	
	$btnRA_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		MSRA.exe /OfferRA $ComputerName
	}	
	
	$btnRestart_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
         $a = new-object -comobject wscript.shell
         $intAnswer = $a.popup( "Are you sure you want to restart $computername", `
0,"Rename Profile",4) 
If ($intAnswer -eq 6) { 
			Invoke-Command  -ComputerName $ComputerName  {shutdown -r -m \\$s -t 15 -f}
} }	

     $btndesk_Click={		
             remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
   		Get-UserName
		Initialize-Listview
		$SBPStatus.Text = "Validating Desktop Paths...."
		#Update-ContextMenu (Get-Variable userSelect*)
		$Properties = $XML.Options.SearchP.Property 
		$Properties | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
		$Info = Get-desk
		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.Properties.(($Col0).ToLower()))
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				[String]$SubItem = $_.Properties.(($Field).ToLower())
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
		$SBPStatus.Text = "Ready"
	}

    $btnProfileUPM1_Click={	
             		   #Load VB module
                   [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                   $a = new-object -comobject wscript.shell
                    Write-Host $txtUser.text
         if($txtUser.text -eq ""){
                       $a.popup( " User textbox cannot empty ")
              }else{	
            remove-ContextMenu(get-Variable Userselect*) 
            remove-ContextMenu(get-Variable GPSelect*) 
   		Get-UserName
		Initialize-Listview
		$SBPStatus.Text = "Validating UPM Profile Paths...."
		#Update-ContextMenu (Get-Variable userSelect*)
		$Properties = $XML.Options.SearchP.Property 
		$Properties | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
		$Info = Get-RPPathUPM
		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.Properties.(($Col0).ToLower()))
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				[String]$SubItem = $_.Properties.(($Field).ToLower())
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
		$SBPStatus.Text = "Ready"
	}}

    $btnUPMRestore_Click={	             		   #Load VB module
                   [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                   $a = new-object -comobject wscript.shell
                    Write-Host $txtUser.text
         if($txtUser.text -eq ""){
                       $a.popup( " User textbox cannot empty ")
              }else{
	
            remove-ContextMenu(get-Variable Userselect*) 
            remove-ContextMenu(get-Variable GPSelect*) 
   		Get-UserName
		Initialize-Listview
		$SBPStatus.Text = "Validating Restore Paths...."
		#Update-ContextMenu (Get-Variable userSelect*)
		$Properties = $XML.Options.SearchP.Property 
		$Properties | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
		$Info = Get-RPPathRestore
		$Info | %{
			$Item = New-Object System.Windows.Forms.ListViewItem($_.Properties.(($Col0).ToLower()))
			ForEach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
				$Field = $Col.Text
				[String]$SubItem = $_.Properties.(($Field).ToLower())
				if($SubItem -ne $null){$Item.SubItems.Add($SubItem)}
				else{$Item.SubItems.Add("")}
			}
			$lvMain.Items.Add($Item)
		}
		$SBPStatus.Text = "Ready"
	} 	}		

	$btnServices_Click={
        remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
		Get-ComputerName
		Services.msc /Computer:$ComputerName
	}	
	
	$btnDash_Click={
		remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
        $URL = "http://dc1-p-a-ctxdir/Director/"
		Start $URL
	}

	$btnPWSync_Click={
		remove-ContextMenu(get-Variable Userselect*) 
        remove-ContextMenu(get-Variable GPSelect*) 
        $URL = "http://corp-w-ws01:8585/SPOP4HD.aspx"
		Start $URL
	}

	$menuViewUPMT_Click={
		$URL = ".\CitrixUPMLogParser.exe"
		Start $URL
	}

	$menuHelpAbout_Click={
		$URL = ".\Help.htm"
		Start $URL
	}
	
	function Get-RPADComputer{
		$Properties = $XML.Options.Search.Property
		if($ComputerName -match "."){$ComputerName = $ComputerName.Split('.')[0]}
		$searcher=[adsisearcher]"(&(objectClass=computer)(name=$ComputerName*))"
		$searcher.PropertiesToLoad.AddRange($Properties) 
		$searcher.SearchRoot=$searchRoot 
		$searcher.FindAll()	
	}

 function Get-RPPathUPM{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $a = new-object -comobject wscript.shell
    $path = "\\dfs-server\Profiles\$username"
    $path1 = "\\dfs-server\Profiles\$username"
    $pathold = "\\dfs-server\Profiles\$username.old"
if(!(Test-Path  $path))
  {

      $a.popup("Profile not found for $username ")
}
else
  {

$intAnswer = $a.popup( "Are you sure you want to rename user porfile $username", `
0,"Rename Profile",4) 
If ($intAnswer -eq 6) { 
    $a.popup("Comfirm in Citrix Director that following user $username is logoff, NOT DISCONNECTED") 
  if(Test-Path  $pathold){

       rm -r $pathold  -Force 

}

    Rename-Item \\dfs-server\Profiles\$username  \\dfs-server1\Profiles\$username.old -force
    Rename-Item \\dfs-server\Profiles\$username  \\dfs-server2\Profiles\$username.old -force
    Rename-Item \\dfs-server1\Profiles\$username  \\dfs-server1\Profiles\$username.old -force
    Rename-Item \\dfs-server2\Profiles\$username  \\dfs-server2\Profiles\$username.old -force
    $a.popup("Profile Rename \\\\dfs-server\Profiles\$username.old ") 

}
if(Test-Path  $path){
    Rename-Item \\dfs-server\Profiles\$username  \\dfs-server1\Profiles\$username.old -force
    Rename-Item \\dfs-server\Profiles\$username  \\dfs-server2\Profiles\$username.old -force
    Rename-Item \\dfs-server1\Profiles\$username  \\dfs-server1\Profiles\$username.old -force
    Rename-Item \\dfs-server2\Profiles\$username  \\dfs-server2\Profiles\$username.old -force
    $a.popup("Profile Rename \\\\dfs-server\Profiles\$username.old ") }
if(Test-Path  $path1){

    Rename-Item \\dfs-server\Profiles\$username  \\dfs-server1\Profiles\$username.old -force
    Rename-Item \\dfs-server\Profiles\$username  \\dfs-server2\Profiles\$username.old -force
    Rename-Item \\dfs-server1\Profiles\$username  \\dfs-server1\Profiles\$username.old -force
    Rename-Item \\dfs-server2\Profiles\$username  \\dfs-server2\Profiles\$username.old -force
    $a.popup("Profile Rename \\\\dfs-server\Profiles\$username.old ") }    
    
    $a.popup("Have the user login now to create a new profile for $username")  

} }

function Get-RPPathRestore{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $a = new-object -comobject wscript.shell
    $path = "backup$\redir\$username"





   II "\\DFS-Server\backup$\redir\$username" -ErrorAction SilentlyContinue 
   II "\\DFS-Server\backup$\redir\$username" -ErrorAction SilentlyContinue 
   II "\\DFS-Server\backup$\redir\$username" -ErrorAction SilentlyContinue 
   II "\\DFS-Server\backup$\redir\$username" -ErrorAction SilentlyContinue 
   II "\\DFS-Server\redir\$username"

  
    
  }

function Get-desk{
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $a = new-object -comobject wscript.shell
    $path = "\\dfs-server\Profiles\$username"
    $path1 = "\\dfs-server\Profiles\$username\VDI\UPM_Profile\desktop"
    $pathCK1 = "\\dfs-server\Profiles\$username\UPM_Profile\desktop"
    $pathCK2 = "\\dfs-server\Profiles\$username\VDI\UPM_Profile\desktop"
if(!(Test-Path  $path))
  {

      $a.popup("Profile not found for $username ")
}
else
  {

$intAnswer = $a.popup( "Are you sure you want to delete desktop file for $username", `
0,"Rename Profile",4) 
If ($intAnswer -eq 6) { 
    $a.popup("Comfirm in Citrix Director that following user $username is logoff, NOT DISCONNECTED") 

    If(Test-Path  $path){
    Remove-Item \\DC1-P-F-DFS01\Profiles\$username\UPM_Profile\desktop   -force
    Remove-Item \\DC1-P-F-DFS02\Profiles\$username\UPM_Profile\desktop   -force
    Remove-Item \\DC2-P-F-DFS01\Profiles\$username\UPM_Profile\desktop   -force
    Remove-Item \\DC2-P-F-DFS02\Profiles\$username\UPM_Profile\desktop   -force
    Remove-Item \\corp\Profiles\$username\UPM_Profile\desktop   -force
    $a.popup("Desktop File has been remove XenApp Profile! checking VDI profile now") 

}
if(Test-Path  $path1){
    Remove-Item \\DC1-P-F-DFS01\Profiles\$username\VDI\UPM_Profile\desktop   -force
    Remove-Item \\DC1-P-F-DFS02\Profiles\$username\VDI\UPM_Profile\desktop   -force
    Remove-Item \\DC2-P-F-DFS01\Profiles\$username\VDI\UPM_Profile\desktop   -force
    Remove-Item \\DC2-P-F-DFS02\Profiles\$username\VDI\UPM_Profile\desktop   -force
    Remove-Item \\corp\Profiles\$username\VDI\UPM_Profile\desktop   -force
    $a.popup("Desktop File has been remove VDI Profile! ")}
if(Test-Path  $pathck1){

    Remove-Item \\dfs-server1\Profiles\$username\UPM_Profile\desktop   -force
    Remove-Item \\dfs-server2\Profiles\$username\UPM_Profile\desktop   -force
    Remove-Item \\dfs-server1\Profiles\$username\UPM_Profile\desktop   -force
    Remove-Item \\dfs-server2\Profiles\$username\UPM_Profile\desktop   -force
        $a.popup("Checking for desktop file ")}    
if(Test-Path  $pathck2){

    Remove-Item \\dfs-server1\Profiles\$username\VDI\UPM_Profile\desktop   -force
    Remove-Item \\dfs-server2\Profiles\$username\VDI\UPM_Profile\desktop   -force
    Remove-Item \\dfs-server1\Profiles\$username\VDI\UPM_Profile\desktop   -force
    Remove-Item \\dfs-server2\Profiles\$username\UPM_Profile\desktop   -force
        $a.popup("Checking for desktop file ")}    }   

    $a.popup("The Desktop file has been delete. Have the user login now to create a new desktop file for $username")  

} } 

	
	function Get-ColumnIndex{	Get-ChildItem –path
		Param($ColumnName)
		$Script:ColumnIndex = ($lvMain.Columns | ?{$_.Text -eq $ColumnName}).Index
	}
	
	function Update-ContextMenu{
		Param($Vis)
		
		Get-Variable cms* | %{Try{$_.Value.Visible = $False}catch{}}
		$Vis | %{try{$_.Value.Visible = $True}catch{}}
	}

	function remove-ContextMenu{
		Param($Vis)
		
		Get-Variable cms* | %{Try{$_.Value.Visible = $true}catch{}}
		$Vis | %{try{$_.Value.Visible = $false}catch{}}
	}
	
	
	function Initialize-Listview{
		$lvMain.Items.Clear()
		$lvMain.Columns.Clear()
       
	}
	
	function Get-ComputerName{
		if($txtComputer.Text -eq "." -or $txtComputer.Text -eq "localhost" -or $txtComputer.Text -eq "" -or $txtComputer.Text -eq $null){$txtComputer.Text = hostname}
		$Script:ComputerName = $txtComputer.Text
		Start-Sleep -Milliseconds 200
	}
	
    function Get-UserName{
		if($txtUser.Text -eq "." -or $txtuser.Text -eq "localhost" -or $txtuser.Text -eq "" -or $txtuser.Text -eq $null){$txtuser.Text = username}
		$Script:UserName = $txtUser.Text
		Start-Sleep -Milliseconds 200
	}
	function Add-Column{
		Param([String]$Column)
		Write-Verbose "Adding $Column from XML file"
		$lvMain.Columns.Add($Column)
	}
	
	function Resize-Columns{
		Write-Verbose "Resizing columns based on column count"
		$ColWidth = (($lvMain.Width / ($lvMain.Columns).Count) - 11)
		$lvMain.Columns | %{$_.Width = $ColWidth}
	}
	
	function Remove-SelectedItems{
		$lvMain.SelectedItems | %{$lvMain.Items.RemoveAt($_.Index)}
	}
	
	function Set-FormTitle{
		$formMain.Text = $XML.Options.Product + " v" + $XML.Options.Version + " - Connected to " + $Domain	
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formMain.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
            $btnNS.remove_Click($btnNS_Click)
            $btnChrome.remove_Click($btnprt_Click)
            $btnHung.remove_Click($btnHung_Click)
            $btnsync.remove_Click($btnsync_Click)
            $btnSecPRT.remove_Click($btnsync_Click)
            $btnSCCMT.remove_Click($btnSCCMT_Click)
            $btngpo.remove_Click($btngpo_Click)
            $btnappv.remove_Click($btnappv_Click)
            $btnProfileupm1.remove_Click($btnProfileupm1_Click)
            $btnUPMRestore.remove_Click($btnUPMRestore_Click)		            			
            $btnRestart.remove_Click($btnRestart_Click)
			$btnRA.remove_Click($btnRA_Click)
			$btnRDP.remove_Click($btnRDP_Click)
            $btnprt.remove_Click($btnprt_Click)
			$btnDSA.remove_Click($btnDSA_Click)
			$btndash.remove_Click($btnDash_Click)
			$btnPWSync.remove_Click($btnPWSync_Click)
			$btnProcesses.remove_Click($btnProcesses_Click)
			$btnEventVwr.remove_Click($btnEventVwr_Click)
			$btnViewUser.remove_Click($btnViewUser_Click)

			$menuViewUPMt.remove_Click($menuViewUPMT_Click)
			$btnServices.remove_Click($btnServices_Click)
			$btnSystemInfo.remove_Click($btnSystemInfo_Click)
			$btnSearch.remove_Click($btnSearch_Click)
			$formMain.remove_Load($formMain_Load)
			$menuFileConnect.remove_Click($menuFileConnect_Click)
			$menuFileExit.remove_Click($menuFileExit_Click)
            $cmsAppUninstall.remove_Click($cmsAppUninstall_Click)
			$cmsProcEnd.remove_Click($cmsProcEnd_Click)
			$cmsSelect.remove_Click($cmsSelect_Click)
			$userSelect.remove_Click($userSelect_Click)
			$userSelectunlock.remove_Click($userSelectunlock_Click) 
			$userSelectGPM.remove_Click($userSelectGPM_Click)
			$GPSelectadd.remove_Click($GPSelectadd_Click)  
			$GPSelectremove.remove_Click($GPSelectremove_Click)     
			$cmsSelogoffuser.remove_Click($cmsSelogoffuser_Click) 
			$menuHelpAbout.remove_Click($menuHelpAbout_Click)
			$formMain.remove_Load($Form_StateCorrection_Load)
			$formMain.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# formMain
	#
	$formMain.Controls.Add($groupTools)
	$formMain.Controls.Add($groupInfo)
	$formMain.Controls.Add($lvMain)
	$formMain.Controls.Add($btnSearch)
	$formMain.Controls.Add($btnSearchuser)
	$formMain.Controls.Add($txtComputer)
    $formMain.Controls.Add($txtUser)
	$formMain.Controls.Add($SB)
	$formMain.Controls.Add($menu)
	$formMain.ClientSize = '980, 710'
	$formMain.MainMenuStrip = $menu
	$formMain.Name = "formMain"
	$formMain.StartPosition = 'CenterScreen'
	$formMain.Text = "OD System Admin Tools v3.0"
	$formMain.add_Load($formMain_Load)
	#
	# groupTools
	#
	$groupTools.Controls.Add($btnChrome)
	$groupTools.Controls.Add($btnNS)
	$groupTools.Controls.Add($btnhung)
	$groupTools.Controls.Add($btnsync)
	$groupTools.Controls.Add($btnSCCMT)	
	$groupTools.Controls.Add($btnappv)
	$groupTools.Controls.Add($btnprofileUPM1)
	$groupTools.Controls.Add($btnprt)
	$groupTools.Controls.Add($btnRestart)
	$groupTools.Controls.Add($btndash)
	$groupTools.Controls.Add($btnPWSync)
	$groupTools.Controls.Add($btnSecPrt)
	$groupTools.Controls.Add($btnRA)
	$groupTools.Controls.Add($btnRDP)
	$groupTools.Controls.Add($btnUPMRestore)
	$groupTools.Location = '10, 360'
	$groupTools.Name = "groupTools"
	$groupTools.Size = '326, 300'
	$groupTools.TabIndex = 8
	$groupTools.TabStop = $False
	$groupTools.Text = "Tools"
	#
	# btnRestart
	#
	$btnRestart.Location = '9, 115'
	$btnRestart.Name = "btnRestart"
	$btnRestart.Size = '110, 25'
	$btnRestart.TabIndex = 12
	$btnRestart.Text = "Restart Computer"
	$btnRestart.UseVisualStyleBackColor = $True
	$btnRestart.add_Click($btnRestart_Click)
    #
	# btnprofileUPM1
	#
	$btnprofileupm1.Location = '9, 50'
	$btnprofileupm1.Name = "btnprofileUPM1"
	$btnprofileupm1.Size = '110, 25'
	$btnprofileupm1.TabIndex = 10
	$btnprofileupm1.Text = "UPM Profile Reset"
	$btnprofileupm1.UseVisualStyleBackColor = $True
	$btnprofileupm1.add_Click($btnprofileUPM1_Click)
	#
	# btnprt
	#
	$btnprt.Location = '9, 140'
	$btnprt.Name = "btnprt"
	$btnprt.Size = '110, 35'
	$btnprt.TabIndex = 10
	$btnprt.Text = "Restart Print Spooler"
	$btnprt.UseVisualStyleBackColor = $True
	$btnprt.add_Click($btnprt_Click)
	# btnUPMRestore
	#
	$btnUPMRestore.Location = '9, 80'
	$btnUPMRestore.Name = "btnUPMRestore"
	$btnUPMRestore.Size = '110, 35'
	$btnUPMRestore.TabIndex = 7
	$btnUPMRestore.Text = "Restore Proifle Files"
	$btnUPMRestore.UseVisualStyleBackColor = $True
	$btnUPMRestore.add_Click($btnUPMRestore_Click)
	#
	# btndash
	#
	$btndash.Location = '9, 25'
	$btndash.Name = "btndash"
	$btndash.Size = '110, 25'
	$btndash.TabIndex = 10
	$btndash.Text = "Citrix Director"
	$btndash.UseVisualStyleBackColor = $True
	$btndash.add_Click($btndash_Click)
	#
	# btnApp-v
	#
	$btnappv.Location = '135, 25'
	$btnappv.Name = "btnappv"
	$btnappv.Size = '110, 25'
	$btnappv.TabIndex = 10
	$btnappv.Text = "App-V Packages"
	$btnappv.UseVisualStyleBackColor = $True
	$btnappv.add_Click($btnappv_Click)
	#
	# $btnSCCMT
	#
	$btnSCCMT.Location = '135, 55'
	$btnSCCMT.Name = "btnSCCMT"
	$btnSCCMT.Size = '110, 25'
	$btnSCCMT.TabIndex = 10
	$btnSCCMT.Text = "SCCM Tool"
	$btnSCCMT.UseVisualStyleBackColor = $True
	$btnSCCMT.add_Click($btnSCCMT_Click)
	#
	# $btnSync
	#
	$btnSync.Location = '135, 80'
	$btnSync.Name = "btnSync"
	$btnSync.Size = '110, 25'
	$btnSync.TabIndex = 10
	$btnSync.Text = "Sync Desktop"
	$btnSync.UseVisualStyleBackColor = $True
	$btnSync.add_Click($btnSync_Click)
	#
	# btnRA
	#
	$btnRA.Location = '135, 112'
	$btnRA.Name = "btnRA"
	$btnRA.Size = '110, 25'
	$btnRA.TabIndex = 9
	$btnRA.Text = "Remote Assistance"
	$btnRA.UseVisualStyleBackColor = $True
	$btnRA.add_Click($btnRA_Click)
	#
	# btnHung
	#
	$btnHung.Location = '135, 142'
	$btnHung.Name = "btnhung"
	$btnHung.Size = '110, 35'
	$btnHung.TabIndex = 9
	$btnHung.Text = "XenApp Hung Session"
	$btnHung.UseVisualStyleBackColor = $True
	$btnHung.add_Click($btnHung_Click)
	#
	# btnNS
	#
	$btnNS.Location = '135, 182'
	$btnNS.Name = "btnNS"
	$btnNS.Size = '110, 35'
	$btnNS.TabIndex = 9
	$btnNS.Text = "New AS400 Session"
	$btnNS.UseVisualStyleBackColor = $True
	$btnNS.add_Click($btnNS_Click)
	#
	# btnRDP
	#
	$btnRDP.Location = '9, 175'
	$btnRDP.Name = "btnRDP"
	$btnRDP.Size = '110, 25'
	$btnRDP.TabIndex = 8
	$btnRDP.Text = "ICA licening issue"
	$btnRDP.UseVisualStyleBackColor = $True
	$btnRDP.add_Click($btnRDP_Click)
	#
	# btnPWSync
	#
	$btnPWSync.Location = '9, 200'
	$btnPWSync.Name = "btnPWSync"
	$btnPWSync.Size = '110, 35'
	$btnPWSync.TabIndex = 8
	$btnPWSync.Text = "Specops Password Sync History"
	$btnPWSync.UseVisualStyleBackColor = $True
	$btnPWSync.add_Click($btnPWSync_Click)
    #
	# btnprofileUPM1
	#
	 $btnChrome.Location = '9, 240'
	 $btnChrome.Name = "btnChrome"
	 $btnChrome.Size = '110, 35'
	 $btnChrome.TabIndex = 10
	 $btnChrome.Text = "Chrome Fixed"
	 $btnChrome.UseVisualStyleBackColor = $True
	 $btnChrome.add_Click( $btnChrome_Click)
	#
	# btnSecPrt
	#
	$btnSecPrt.Location = '135, 220'
	$btnSecPrt.Name = "btnSecPrt"
	$btnSecPrt.Size = '110, 35'
	$btnSecPrt.TabIndex = 10
	$btnSecPrt.Text = "Secure Print"
	$btnSecPrt.UseVisualStyleBackColor = $True
	$btnSecPrt.add_Click($btnSecPrt_Click)
	#
	# groupInfo
	#
	$groupInfo.Controls.Add($btngpo)
	$groupInfo.Controls.Add($btnViewUser)
	$groupInfo.Controls.Add($btnDSA)
	$groupInfo.Controls.Add($btnProcesses)
	$groupInfo.Controls.Add($btnEventVwr)
	$groupInfo.Controls.Add($btnServices)
	$groupInfo.Controls.Add($btnSystemInfo)
	$groupInfo.Location = '10, 135'
	$groupInfo.Name = "groupInfo"
	$groupInfo.Size = '326, 208'
	$groupInfo.TabIndex = 7
	$groupInfo.TabStop = $False
	$groupInfo.Text = "Information"
	#
	# btngpo
	#
	$btngpo.Location = '135, 105'
	$btngpo.Name = "btngpo"
	$btngpo.Size = '110, 25'
	$btngpo.TabIndex = 10
	$btngpo.Text = "GPO Result"
	$btngpo.UseVisualStyleBackColor = $True
	$btngpo.add_Click($btngpo_Click)
	#
	# btnViewUser
	#
	$btnViewUser.Location = '135,19'
	$btnViewUser.Name = "btnViewUser"
	$btnViewUser.Size = '110, 35'
	$btnViewUser.TabIndex = 7
	$btnViewUser.Text = "Local Users and Groups"
	$btnViewUser.UseVisualStyleBackColor = $True
	$btnViewUser.add_Click($btnViewUser_Click)
	#
	# btnDSA
	#
	$btnDSA.Location = '135, 65'
	$btnDSA.Name = "btnDSA"
	$btnDSA.Size = '110, 30'
	$btnDSA.TabIndex = 7
	$btnDSA.Text = "AD Users and Computers"
	$btnDSA.UseVisualStyleBackColor = $True
	$btnDSA.add_Click($btnDSA_Click)
	#
	# btnProcesses
	#
	$btnProcesses.Location = '9, 105'
	$btnProcesses.Name = "btnProcesses"
	$btnProcesses.Size = '110, 25'
	$btnProcesses.TabIndex = 6
	$btnProcesses.Text = "Processes"
	$btnProcesses.UseVisualStyleBackColor = $True
	$btnProcesses.add_Click($btnProcesses_Click)
	#
	# btnEventVwr
	#
	$btnEventVwr.Location = '9, 78'
	$btnEventVwr.Name = "btnEventVwr"
	$btnEventVwr.Size = '110, 25'
	$btnEventVwr.TabIndex = 5
	$btnEventVwr.Text = "Event Viewer"
	$btnEventVwr.UseVisualStyleBackColor = $True
	$btnEventVwr.add_Click($btnEventVwr_Click)
	#
	# btnservice
	#
	$btnServices.Location = '9, 50'
	$btnServices.Name = "btnServices"
	$btnServices.Size = '110, 25'
	$btnServices.TabIndex = 3
	$btnServices.Text = "Local Services"
	$btnServices.UseVisualStyleBackColor = $True
	$btnServices.add_Click($btnServices_Click)
	#
	# btnSystemInfo
	#
	$btnSystemInfo.Location = '9, 19'
	$btnSystemInfo.Name = "btnSystemInfo"
	$btnSystemInfo.Size = '110, 25'
	$btnSystemInfo.TabIndex = 2
	$btnSystemInfo.Text = "System Info"
	$btnSystemInfo.UseVisualStyleBackColor = $True
	$btnSystemInfo.add_Click($btnSystemInfo_Click)
	#
	# lvMain
	#
	$lvMain.Anchor = 'Top, Bottom, Left, Right'
	$lvMain.ContextMenuStrip = $contextMenu
	$lvMain.FullRowSelect = $True
	$lvMain.GridLines = $True
	$lvMain.Location = '342, 28'
	$lvMain.Name = "lvMain"
	$lvMain.Size = '630, 590'
	$lvMain.TabIndex = 13
	$lvMain.UseCompatibleStateImageBehavior = $False
	$lvMain.View = 'Details'
	#
	# btnSearch
	#
	$btnSearch.Location = '19, 52'
	$btnSearch.Name = "btnSearch"
	$btnSearch.Size = '110, 25'
	$btnSearch.TabIndex = 1
	$btnSearch.Text = "Search for PC"
	$btnSearch.UseVisualStyleBackColor = $True
	$btnSearch.add_Click($btnSearch_Click)
	#
	# txtComputer
	#
	$txtComputer.Location = '10, 28'
	$txtComputer.Name = "txtComputer"
	$txtComputer.Size = '126, 20'
	$txtComputer.TabIndex = 0
	#
    # txtUser
	#
	$txtUser.Location = '10, 80'
	$txtUser.Name = "txtUser"
	$txtUser.Size = '126, 20'
	$txtUser.TabIndex = 0
    #
	# btnSearchUser
	#
	$btnSearchUser.Location = '19, 102'
	$btnSearchUser.Name = "btnSearchUser"
	$btnSearchUser.Size = '110, 25'
	$btnSearchUser.TabIndex = 1
	$btnSearchUser.Text = "Search for User"
	$btnSearchUser.UseVisualStyleBackColor = $True
	$btnSearchUser.add_Click($btnSearchUser_Click)
    #
	# SB
	#
	$SB.Anchor = 'Bottom, Left, Right'
	$SB.Dock = 'None'
	$SB.Location = '0, 660'
	$SB.Name = "SB"
	[void]$SB.Panels.Add($SBPBlog)
	[void]$SB.Panels.Add($SBPStatus)
	$SB.ShowPanels = $True
	$SB.Size = '980, 22'
	$SB.TabIndex = 1
	$SB.Text = "Ready"
	#
	# menu
	#
	[void]$menu.Items.Add($menuFile)
	[void]$menu.Items.Add($menuView)
	[void]$menu.Items.Add($menuHelp)
	$menu.Location = '0, 0'
	$menu.Name = "menu"
	$menu.Size = '780, 24'
	$menu.TabIndex = 0
	$menu.Text = "menuMain"
	#
	# menuFile
	#
	[void]$menuFile.DropDownItems.Add($menuFileExit)
	$menuFile.Name = "menuFile"
	$menuFile.Size = '37, 20'
	$menuFile.Text = "File"
	#
	# menuFileExit
	#
	$menuFileExit.Name = "menuFileExit"	$menuFileExit.Size = '186, 22'
	$menuFileExit.Text = "Exit"
	$menuFileExit.add_Click($menuFileExit_Click)
	#
	# menuView
	#
	[void]$menuView.DropDownItems.Add($menuViewUPMT)
	$menuView.Name = "menuView"
	$menuView.Size = '44, 20'
	$menuView.Text = "Links"
    #
	$menuViewUPMT.Name = "menuViewUPMT"
	$menuViewUPMT.Size = '152, 22'
	$menuViewUPMT.Text = "Citrix UPM Log Parser"
	$menuViewUPMT.add_Click($menuViewUPMT_Click)
	# contextMenu
	#
	[void]$contextMenu.Items.Add($userSelectunfind)
	[void]$contextMenu.Items.Add($GPSelectadd)
	[void]$contextMenu.Items.Add($GPSelectremove)
	[void]$contextMenu.Items.Add($userSelectgpm)
	[void]$contextMenu.Items.Add($userSelectUnlock)
	[void]$contextMenu.Items.Add($cmsProcEnd)
	[void]$contextMenu.Items.Add($cmsSelect)
	[void]$contextMenu.Items.Add($userSelect) 
	[void]$contextMenu.Items.Add($cmsSelogoffuser)
	[void]$contextMenu.Items.Add($cmsAppUninstall)
	$contextMenu.Name = "contextMenu"
	$contextMenu.Size = '188, 114'

	#
	# $userSelectunfind
	#
	$userSelectunfind.Name = "userSelectunfind"
	$userSelectunfind.Size = '187, 22'
	$userSelectunfind.Text = "What Device is Locking Out an user"
	$userSelectunfind.Visible = $False
	$userSelectunfind.add_Click($userSelectunfind_Click)
	#
	#
	# $GPSelectremove
	#
	$GPSelectremove.Name = "GPSelectremove"
	$GPSelectremove.Size = '187, 22'
	$GPSelectremove.Text = "Remove $username from a Group"
	$GPSelectremove.Visible = $False
	$GPSelectremove.add_Click($GPSelectremove_Click)
	#
	# $GPSelectadd
	#
	$GPSelectadd.Name = "GPSelectadd"
	$GPSelectadd.Size = '187, 22'
	$GPSelectadd.Text = "Add $username to a Group"
	$GPSelectadd.Visible = $False
	$GPSelectadd.add_Click($GPSelectadd_Click)
	#
	# userSelectGPM
	#
	$userSelectGPM.Name = "userSelectGPM"
	$userSelectGPM.Size = '187, 22'
	$userSelectGPM.Text = "Group Membership"
	$userSelectGPM.Visible = $False
	$userSelectGPM.add_Click($userSelectGPM_Click)
	#
	# userSelectUnlock
	#
	$userSelectUnlock.Name = "userSelectUnlock"
	$userSelectUnlock.Size = '187, 22'
	$userSelectUnlock.Text = "Unlock"
	$userSelectUnlock.Visible = $False
	$userSelectUnlock.add_Click($userSelectUnlock_Click)
	#
	# cmsAppUninstall
	#
	$cmsAppUninstall.Name = "cmsAppUninstall"
	$cmsAppUninstall.Size = '187, 22'
	$cmsAppUninstall.Text = "Reinstall App-v Package"
	$cmsAppUninstall.Visible = $False
	$cmsAppUninstall.add_Click($cmsAppUninstall_Click)
	#
	# cmsProcEnd
	#
	$cmsProcEnd.Name = "cmsProcEnd"
	$cmsProcEnd.Size = '187, 22'
	$cmsProcEnd.Text = "End Process"
	$cmsProcEnd.Visible = $False
	$cmsProcEnd.add_Click($cmsProcEnd_Click)
	#
	# cmsSelect
	#
	$cmsSelect.Name = "cmsSelect"
	$cmsSelect.Size = '187, 22'
	$cmsSelect.Text = "Select Computer"
	$cmsSelect.Visible = $False
	$cmsSelect.add_Click($cmsSelect_Click)
	#
	# userSelect
	#
	$userSelect.Name = "userSelect"
	$userSelect.Size = '187, 22'
	$userSelect.Text = "Select User"
	$userSelect.Visible = $False
	$userSelect.add_Click($userSelect_Click)
	#
	# menuHelp
	#
	[void]$menuHelp.DropDownItems.Add($menuHelpAbout)
	$menuHelp.Name = "menuHelp"
	$menuHelp.Size = '44, 20'
	$menuHelp.Text = "Help"
	#
	# menuHelpAbout
	#
	$menuHelpAbout.Name = "menuHelpAbout"
	$menuHelpAbout.Size = '152, 22'
	$menuHelpAbout.Text = "Help file"
	$menuHelpAbout.add_Click($menuHelpAbout_Click)
	#
	# SBPStatus
	#
	$SBPStatus.AutoSize = 'Spring'
	$SBPStatus.Name = "SBPStatus"
	$SBPStatus.Text = "Ready"
	$SBPStatus.Width = 60
	#
	# SBPBlog
	#
	$SBPBlog.Alignment = 'Center'
	$SBPBlog.Name = "SBPBlog"
	$SBPBlog.Text = "http://portal.company.com/"
	$SBPBlog.Width = 160
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formMain.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formMain.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formMain.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formMain.ShowDialog()

} #End Function

#Call OnApplicationLoad to initialize
if((OnApplicationLoad) -eq $true)
{
	#Call the form
	Call-odhd_pff | Out-Null
	#Perform cleanup
	OnApplicationExit
$now=Get-Date -format "dd-MMM-yyyy HH:mm"
Write-output  "$env:UserName ,  $now , Stop"| Out-File -Append  .\logs\access-logs.csv
}
