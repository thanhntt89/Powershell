<# Setting: Message box #>
Add-Type -AssemblyName PresentationCore, PresentationFramework

<# Setting: Input box #>
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

########### CALL FUNCTION #####################################################################
#Call: INS.PS1 para1 para2
#Para1: ModuleType(Value[1,2] 1: CollectionNet 2: CollectionNet_Z)
#Para2: MethodType(Value[null,(int)]) null or NOT equal[1;2]: Default 1: Method1 2: Methode2

##### WSHスクリプト指定 ########################################################################
$formWaiting = New-Object system.Windows.Forms.Form
$formInputNumber = New-Object system.Windows.Forms.Form
$formSelection = New-Object system.Windows.Forms.Form
$WSC = new-object -comobject wscript.shell

######################################################################## WSHスクリプト指定 #####

####################################BEGIN_VARIABLE############################################
#Get input parameter 
[Int]$global:METHOD_TYPE = $args[0]
$VER = "02.2021"

$ACOMP_SWITCH = 0

$CURRENT_LOCATION = "C:\cms5"
#INIT_EEROR_SET
$global:COMP_ERR = -1
$global:IsStopThread = $false;

#####################BEGIN_INS1############################

#フロッピードライブ名
$FD_DRV = "C:"

#CAの出力ディレクトリ
$DB_DIR = "C:\cms5\DB"

#バッチ内ＴＭＰファイル、作成ディレクトリ
$TMP_DIR = "C:\CMS5\DATA"
$EXE_DIR = "C:\CMS5\XINGTOOL"

#バッチ内ＴＭＰファイル
$TMP_NAME = "SELECT.TMP"

#ツール類
$SMF_TO_RCP_EXE = "C:\X_Tools\SMF to RCP\SMFtoRCP.exe"
$SMF_TO_RCP_LOG = "C:\X_Tools\SMF to RCP\SMFtoRCP.log"
$RCP_TO_SMF_EXE = "C:\X_Tools\RCP to SMF\RCPtoSMF.exe"
$RCP_TO_SMF_LOG = "C:\X_Tools\RCP to SMF\RCPtoSMF.log"

##TOOLS_FIXED
$UNLHA_EXE_PATH = "$EXE_DIR\unlha.exe"
$GET_INTRO_TIME_FROM_RCP_EXE_PATH = "$EXE_DIR\GetIntroTimeFromRcp.exe"
$UNION_CSV_EXE_PATH = "$EXE_DIR\UNION_CSV.exe"
$GET_DBN2_EXE_PATH = "$EXE_DIR\getDBN2.exe"
$RCP_CHKS_EXE_PATH = "$EXE_DIR\RCPChks.exe"
#$PUT_JD5_EXE_PATH = "$EXE_DIR\putJD5.exe"
$JD5_PAR_SET_JD6_PATH = "$EXE_DIR\JD5ParSetJD6.exe"
$WBLHA_EXE_PATH = "$EXE_DIR\WBLHA.exe"
$REPLACE_CHAR_DATA_EXE_PATH = "$EXE_DIR\Replace_CharData.exe"
$PLA_COL_PALLET_CHG_EXE_PATH = "$EXE_DIR\PLAColPalletChg.exe"
$JD7_DVTON_EXE_PATH = "$EXE_DIR\JD7_DVTOn.exe"
$UGA_TITLE_WORDS_TO_BCHG_EXE_PATH = "$EXE_DIR\UGATitleWordStoBChg.exe"
$COMPARE_JD7MID_EXE_PATH = "$EXE_DIR\CompareJD7MID.exe"
$INPUR_DDB_EXE_PATH = "$EXE_DIR\InputDDB.exe"
$JD5_CHECK_EXE_PATH = "$EXE_DIR\JD5check.exe"

#INIT_Global variable 
#6桁ディレクトリ作成変数
$global:re = ""
$global:KANRI_D = ""
$global:BK_DRV = ""
$global:RCP_DIR = ""
$global:AAC_DIR = ""
$global:KEN_DIR = ""
$global:ZEUS_DIR = ""
$global:FLG_DIR = ""
$global:FLG_DIR = ""
$global:BK_DIR = ""
$global:SLC = ""
$global:file_name = ""
$global:PSLC = ""
$global:ERR_FNAME = ""
$global:FN = ""
$global:FN_HEAD = ""
$global:RN = ""
$global:EXTENSION = ""
$global:DRV = ""
$global:R88_N = ""
$global:R88_REV = ""
$global:REV = ""
$global:R55_N = ""
$global:R55_REV = ""
$global:RAA_N = ""
$global:RAA_REV = ""
$global:RBB_N = ""
$global:RBB_REV = ""
$global:RCC_N = ""
$global:RCC_REV = ""
$global:INTRO = ""
$global:FLG_NAME = ""
$global:sh = ""

#####################END_INS1############################

####################################BEGIN_MODULE_CONFIG########################################
#Init INS

if ($global:METHOD_TYPE -eq 0) {
    $global:KANRI_D = "\\JS-server\kanri"
    $global:BK_DRV = "\\JS-server\MAKEDATA" #"C:\MAKEDATA"#
    $global:RCP_DIR = "\\JS-server\kanri\RCP"
    $global:AAC_DIR = "\\kc-server\8Ch_Data" 
    $global:KEN_DIR = "\\kc-server\ken_data"
    $global:ZEUS_DIR = "\\kc-server\zeus_data"
}
elseif ($global:METHOD_TYPE -eq 1) {
    $global:KANRI_D = "C:\BACK_UP"
    $global:BK_DRV = "A:"
}
elseif ($global:METHOD_TYPE -eq 2) {
    $global:KANRI_D = "S:"
    $global:BK_DRV = "S:"
    $global:FLG_DIR = "S:\FLG"
    $global:RCP_DIR = "S:\RCP"
    $global:AAC_DIR = "S:\SCREC\8ch_data"
}
elseif ($global:METHOD_TYPE -eq 3) {
    $global:KANRI_D = "C:\back_up"
    $global:BK_DRV = "C:\back_up"
    $global:FLG_DIR = "C:\back_up\FLG"
    $global:RCP_DIR = "C:\back_up\RCP"
}
elseif ($global:METHOD_TYPE -eq 4) {
    $global:KANRI_D = "S:"
    $global:BK_DRV = "S:"
    $global:FLG_DIR = "S:\FLG"
    $global:RCP_DIR = "S:\RCP"
    $global:AAC_DIR = "S:\SCREC\8ch_data"
}
elseif ($global:METHOD_TYPE -eq 5) {
    $global:KANRI_D = "\\KC-server\テロップグレードアップ\DATA"
    $global:BK_DRV = "\\KC-server\テロップグレードアップ\DATA" #"C:\DATA"#
    $global:FLG_DIR = "\\KC-server\テロップグレードアップ\DATA\FLG"
    $global:RCP_DIR = "\\KC-server\テロップグレードアップ\DATA\RCP"
    $global:AAC_DIR = "\\KC-server\テロップグレードアップ\DATA\SCREC\8ch_data"
}
elseif ($global:METHOD_TYPE -eq 6) {
    $global:KANRI_D = "\\kc-server\orchestra\Overseas\MASTER\F1_revised"
    $global:BK_DRV = "\\kc-server\orchestra\Overseas\MASTER\F1_revised" #"C:\F1_revised" #
    $global:FLG_DIR = "\\kc-server\orchestra\Overseas\MASTER\F1_revised\FLG"
    $global:RCP_DIR = "\\kc-server\orchestra\Overseas\MASTER\F1_revised\RCP"
    $global:AAC_DIR = "\\kc-server\orchestra\Overseas\MASTER\F1_revised\SCREC\8ch_data"
}
else {
    NO_PRAM
}

####################################END_MODULE_CONFIG##########################################

#Define array to loop
$ARRAYS = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

####################################END_VARIABLE##############################################

#Get input paramter
$global:input = $args 

function ValidInput {
    #Not input parameter
    if ($global:input.Count -ne 1 -or $global:input -gt 6) {
        $WSC.popup("You must input one parameter with value:`n`nXING In-house Param:[0,1,3,5,6]`n`nInspection software house Param:[4]`n`nDomestic production software house Param:[2]`n`nOverseas production software house Param:[2]", 0, "`t<< ERROR >>", 0)
     
        if ($formWaiting.Visible -eq $true) {
            $formWaiting.Close()
        }

        exit
    }
   
    ##Valid param value
    if ((ValidParamterInput($global:METHOD_TYPE)) -eq $false) {
        $WSC.popup("Parmeter must in range [0,1,2,3,4,5,6]", 0, "`t<< ERROR >>", 0)
        if ($formWaiting.Visible -eq $true) {
            $formWaiting.Close()
        }
        exit      
    }        
}

####################################BEGIN_CHECK_INPUT_PARA########################################

####################################END_CHECK_INPUT_PARA##########################################

#########################################BEGIN_UTILS#############################################

function CheckDefaultFolder {
      
    if ((CheckExistPath($TMP_DIR)) -ne $true) {
        $WSC.popup("Folder $TMP_DIR not exist!", 0, "`t<< ERROR >>", 0)       
        THE_END        
    }

    if ((CheckExistPath($CURRENT_LOCATION)) -ne $true) {
        $WSC.popup("Folder $CURRENT_LOCATION not exist!", 0, "`t<< ERROR >>", 0)       
        THE_END
    }

    if ((CheckExistPath($EXE_DIR)) -ne $true) {
        $WSC.popup("Folder $EXE_DIR not exist!", 0, "`t<< ERROR >>", 0)      
        THE_END
    }

    if ((CheckExistPath($DB_DIR)) -ne $true) {
        $WSC.popup("Folder $DB_DIR not exist!", 0, "`t<< ERROR >>", 0)
        THE_END
    }

    if ((CheckExistPath($FD_DRV)) -ne $true) {
        $WSC.popup("Folder $FD_DRV not exist!", 0, "`t<< ERROR >>", 0)
        THE_END
    }
}

function ValidParamterInput {
    param (
        [Parameter(Mandatory = $True)]
        [ValidateRange(0, 6)]
        [int] $paramter        
    )

    if ($paramter -eq 0) {
        return $True
    }
    if ($paramter -eq 1) {
        return $True
    }
    if ($paramter -eq 2) {
        return $True
    }
    if ($paramter -eq 3) {
        return $True
    }
    if ($paramter -eq 4) {
        return $True
    }
    if ($paramter -eq 5) {
        return $True
    }
    if ($paramter -eq 6) {
        return $True
    }

    return $False
}

#Get folder number
function GetFolderNumber {
    param (
        [Parameter(Mandatory = $True)]
        [String] $folderIndex        
    )

    return  [String]([Math]::floor(([int]$folderIndex - 1) / 100 + 1) * 100)    
}

#Show form select from list
function ShowSelectionForm {
    #$form = New-Object System.Windows.Forms.Form
    $formSelection.Text = 'Get PTN File'
    $formSelection.MaximizeBox = $False
    $formSelection.Size = New-Object System.Drawing.Size(390, 380)
    $formSelection.StartPosition = 'CenterScreen'
    $formSelection.FormBorderStyle = 'FixedDialog'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(295, 129)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formSelection.AcceptButton = $okButton
    $formSelection.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(295, 158)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formSelection.CancelButton = $cancelButton
    $formSelection.Controls.Add($cancelButton)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(12, 12)
    $label.Size = New-Object System.Drawing.Size(260, 25)
    $label.Text = "Get PTN File for Title Data.`nPlease Select Number!:"
    $formSelection.Controls.Add($label)
    
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(12, 40)
    $listBox.Size = New-Object System.Drawing.Size(280, 290)
    $listBox.Height = 290
    
    [void] $listBox.Items.Add('[01].RED-1.PTN')
    [void] $listBox.Items.Add('[02].RED-3.PTN')
    [void] $listBox.Items.Add('[03].GREEN-1.PTN')
    [void] $listBox.Items.Add('[04].GREEN-3.PTN')
    [void] $listBox.Items.Add('[05].BLUE-1.PTN')
    [void] $listBox.Items.Add('[06].BLUE-3.PTN')
    [void] $listBox.Items.Add('[07].POP-B.PTN')
    [void] $listBox.Items.Add('[08].POP-G.PTN')
    [void] $listBox.Items.Add('[09].POP-M.PTN')
    [void] $listBox.Items.Add('[10].POP-M2.PTN')
    [void] $listBox.Items.Add('[11].POP-P.PTN')
    $formSelection.Controls.Add($listBox)
    
    $formSelection.Topmost = $true
    
    $formInputNumber.Add_Shown( { $formSelection.Activate() })
    $formInputNumber.Add_Shown( { $listBox.Select() })

    $result = $formSelection.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $itemSelected = $listBox.SelectedIndex      
    }
    else {
        $itemSelected = -1
    }
    return  $itemSelected
}

<#Save file info
SaveFileInfo sourceFile.exttion desFile.extension
#>
function SaveFileInfo {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $filePathSource,
        [Parameter(Mandatory = $True)]
        [String] $filePathDes
    )
    
    if ((CheckExistPath($filePathSource)) -eq $True) {
        Get-ChildItem $filePathSource -recurse -name | Set-Content $filePathDes
    }
}

<#Get subs array in arrays#>
function GetSubArrayInArrays {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $contents,
        [Parameter(Mandatory = $True)]
        [char] $char,
        [Parameter(Mandatory = $True)]
        [Int] $startIndex,  
        [Parameter(Mandatory = $True)] 
        [Int] $numberItem     
    )

    $array = $contents.Split("$char")
   
    
    for ($i = $startIndex; $i -lt $array.Count; $i++) {
        if ($i -gt ($startIndex + $numberItem)) {
            break
        }
        $data += $array[$i];
    }

    return $data 
}

#Read all line in file
function ReadAllText {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $filePath
    )
    $contents = ""

    $newstreamreader = New-Object System.IO.StreamReader($filePath)      
    while ($null -ne ($readeachline = $newstreamreader.ReadLine())) {           
        $contents += "$readeachline`n"
    }
    $newstreamreader.Dispose()

    return $contents 
}

<#Rename file#>
function ReName {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $sourceFilePath,
        [Parameter(Mandatory = $True)]
        [String] $desFilePath     
    )
    if ((CheckExistPath($sourceFilePath)) -eq $True) {
        Rename-Item -Path $sourceFilePath -NewName $desFilePath
    }  
}

<#Count file in folder
Count file with extension
#>
function CountFileExtention {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $folderPath,
        [Parameter(Mandatory = $True)]
        [String] $extention
    )
    
    return (Get-ChildItem $folderPath -Recurse -Filter "$extention").Count
}

#Run application with parameter
#Application.exe parameter[]
function StartProcessWait {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $fileExe,
        [Parameter(Mandatory = $True)]
        [String[]] $arg
    )


    if ((CheckExistPath($fileExe)) -eq $True) {
        $params = [system.String]::Join(" ", $arg)  

        Start-Process -Wait $fileExe $params
    }
    else {
        Write-Host "Can't find exe to run EXE_FILE_PATH: $fileExe"
    }  
}

<#Run exe#>
function RUN_EXE {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $fileExe,
        [Parameter(Mandatory = $True)]
        [String[]] $arg
    )
    if ((CheckExistPath($fileExe)) -eq $True) {
        $params = [system.String]::Join(" ", $arg)    
        #Write-Host "RUN_EXE:paramter cmd.exe /c $fileExe $arg"
        cmd.exe /c $fileExe $params
    }
    else {
        Write-Host "Can't find exe to run EXE_FILE_PATH: $fileExe"
    }    
}

<#Copy data#>
function COPY {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $filePathSource,
        [Parameter(Mandatory = $True)]
        [String] $filePathDes
    )
    if ( $global:IsStopThread -eq $true) {
        THE_END
        return
    }
    try {
        Copy-Item -Path $filePathSource -Destination $filePathDes  -Recurse -Force
    }
    catch {           
        Write-Host $Error[0].Exception.Message
    }    
}

<#Reading data in line
Indedex to read
#>
function GetContentByIndex {
    param (
        [Parameter(Mandatory = $True)]
        [String] $filePath,
        [Parameter(Mandatory = $True)]
        [int] $lineIndex
    )
    $data = ""
    try {
        if ((CheckExistPath($filePath)) -eq $True) {
            $data = ((Get-Content -Path $filePath) -as [string[]])[$lineIndex]
        }
    }
    catch {        
    }
    return $data
}

<#
Delete file in folder (folderpath\*.*)
Delete folder include current foler (folderPath\)
#>
function DeleteFileOrFolder {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $itemPath
    )

    if ((CheckExistPath($itemPath)) -eq $True) {
        Remove-Item -Path $itemPath 
    }
}

#Create directory
function MakeDirectory {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $folderPath
    )

    if ((CheckExistPath($folderPath)) -ne $True) {
        mkdir $folderPath
    }   
}

#Write log
function WriteLog {
    param (      
        [Parameter(Mandatory = $True)]
        [String] $contents
    )
    
    WriteToFile $LOG_FILE_PATH $contents
}

<#Save data to file#>
function WriteToFile {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $filePath,
        [Parameter(Mandatory = $True)]
        [String] $contents
    )

    if ((CheckExistPath($filePath)) -eq $True) {
        Add-Content -Path $filePath -Value $contents -Force
    }
    else {
        Set-Content -Path $filePath -Value $contents
    }   
}

<#Check file or folder exist#>
function CheckExistPath {
    Param(
        [Parameter(Mandatory = $True)]
        [String] $checkPath
    )   
    if ( $global:IsStopThread -eq $true) {
        THE_END
        return $false
    }
    if ([String]::IsNullOrWhiteSpace($checkPath)) {
        return $false
    }
    if (-not (Test-Path -Path $checkPath)) {
        return $false
    }
    else {      
        return $True
    }
}

<#Form input number#>
function InputSongNumber {
    Param(       
        [String] $lblText,
        [String] $defaultText = ""
    )  
    
    #$form = New-Object System.Windows.Forms.Form
    $formInputNumber.Text = 'Input Data'
    $formInputNumber.Size = New-Object System.Drawing.Size(570, 300)
    $formInputNumber.StartPosition = 'CenterScreen'
    $formInputNumber.MaximizeBox = $False
    $formInputNumber.FormBorderStyle = 'FixedDialog'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Font = New-Object System.Drawing.Font("Lucida Console", 12, [System.Drawing.FontStyle]::Regular)
    $okButton.Location = New-Object System.Drawing.Point(120, 155)
    $okButton.Size = New-Object System.Drawing.Size(102, 40)
    $okButton.Text = 'OK'   
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formInputNumber.AcceptButton = $okButton
    $formInputNumber.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Font = New-Object System.Drawing.Font("Lucida Console", 12, [System.Drawing.FontStyle]::Regular)
    $cancelButton.Location = New-Object System.Drawing.Point(270, 155)
    $cancelButton.Size = New-Object System.Drawing.Size(150, 40)
    $cancelButton.Text = 'Cancel'   
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel    
    $formInputNumber.Controls.Add($cancelButton)

    $label = New-Object System.Windows.Forms.Label     
    $label.Font = New-Object System.Drawing.Font("Lucida Console", 12, [System.Drawing.FontStyle]::Regular)   
    $label.Location = New-Object System.Drawing.Point(20, 28)
    $label.Size = New-Object System.Drawing.Size(600, 30)
    $label.Text = $lblText
    $formInputNumber.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Font = New-Object System.Drawing.Font("Lucida Console", 20, [System.Drawing.FontStyle]::Regular)
    $textBox.Multiline = $true   
    $textBox.Location = New-Object System.Drawing.Point(25, 75)
    $textBox.Size = New-Object System.Drawing.Size(500, 37)
    $textBox.Text = $defaultText    
    $formInputNumber.Controls.Add($textBox)

    $formInputNumber.Topmost = $true

    $formInputNumber.Add_Shown( { $formInputNumber.Activate() })
    $formInputNumber.Add_Shown( { $textBox.Select() })
    $result = $formInputNumber.ShowDialog()

    $x = ""

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $x = $textBox.Text       
    }
    
    return $x
}

<#Display message info#>
function  ShowMessageInfo {
    param ( 
        [String]$msgTitle,   
        [String]$msgBody      
    ) 
    $msgButton = "OK"
    $msgIcon = "Info"
    $WSC.popup(" $msgBody", 0, "`t<< $msgTitle >>", 0)
    #ShowMessage $msgTitle $msgBody $msgButton $msgIcon  
}

<#Display message error#>
function  ShowMessageError {
    param ( 
        [String]$msgTitle,   
        [String]$msgBody      
    ) 
    $msgButton = "OK"
    $msgIcon = "Error"
    $WSC.popup(" $msgBody", 0, "`t<< $msgTitle >>", 0)

    #ShowMessage $msgTitle $msgBody $msgButton $msgIcon  
}

<#Show message#>
function ShowMessage {
    param (        
        [String]$msgTitle,
        [String]$msgBody,
        [String]$msgButton,
        [String]$msgIcon
    )    
    $iconType = 0
    $btnType = 0
    if ($msgIcon -eq "None") {
        $iconType = 0
    }
    elseif ($msgIcon -eq "Error" -or $msgIcon -eq "Stop") {
        $iconType = 16
    }
    elseif ($msgIcon -eq "Question") {
        $iconType = 32
    }
    elseif ($msgIcon -eq "Warning") {
        $iconType = 48
    }
    elseif ($msgIcon -eq "Info") {
        $iconType = 64
    }

    if ($msgButton -eq "OK") {
        $btnType = 0
    }
    elseif ($msgButton -eq "OKCancel") {
        $btnType = 1
    }
    elseif ($msgButton -eq "YesNoCancel") {
        $btnType = 3
    }
    elseif ($msgButton -eq "YesNo") {
        $btnType = 4
    }

    [System.Windows.MessageBox]::Show($msgBody, $msgTitle, $btnType, $iconType)
}
 
#########################################END_UTILS##############################################

#########################################START_SYSTEM_MESSAGE###################################################

#NO_PRAM
function NO_PRAM {
    $msgTitle = "Parameter Error!!"
    $msgBody = "Usage : INS.PS1 [0..6`n`t0=For running"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END  
}

#CHIRUDA_ERROR
function CHIRUDA_ERROR {
    $msgTitle = "File Name Error!!"
    $msgBody = "RCP File Name Error!!!($global:DRV Folder)`n ex:AB1234~1.R36"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END  
}

#NO_BASEFILE
function NO_BASEFILE {
    $msgTitle = "File Checking Error!!"
    $msgBody = "Not found W or G file of RCP`n`nPlease check in FD"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END      
}

#SADCP_ERR
function SADCP_ERR {
    $msgTitle = "ERROR"
    $msgBody = "BMN File is not able to Get BeMAX Number [C:\cms5\DB]!!"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END         
}

#NO_FILE_ERR
function NO_FILE_ERR {
    $msgTitle = "File Checking Error!!"
    $msgBody = "$global:ERR_FNAME file is not found!!"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END      
}

#NO_TMP_FILE_ERR
function NO_TMP_FILE_ERR {
    $msgTitle = "File Checking Error!!"
    $msgBody = "TMP file is not found!!"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END      
}

#NO_FN_FILE
function NO_FN_FILE {
    $msgTitle = "File Checking Error!!"
    $msgBody = "Not found FN file in FD"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END      
}

#SADCP_ERR
function SADCP_ERR {
    $msgTitle = "ERROR"
    $msgBody = "BMN File is not able to Get BeMAX Number [C:\cms5\DB]!!"
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END  
}

#JD5CHG_ERR
function JD5CHG_ERR {
    $msgTitle = "ERROR"
    $msgBody = "JD5 File Parameter is not able to Change [C:\cms5\data]!!"   
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END  
}

#NODBN_NETWORK
function NODBN_NETWORK {
    $msgTitle = "ERROR"
    $msgBody = "DBN File is Nothing in [S:\\FLG] !!"    
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END  
}

#NOFLG_NETWORK
function NOFLG_NETWORK {
    $msgTitle = "ERROR"
    $msgBody = "FLG File is Nothing in [S:\\FLG] !!"   
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END  
}

#NODBN
function NODBN {
    $msgTitle = "ERROR"
    $msgBody = "DBN File is Nothing in [C:\\cms5\\DB] !!"   
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END  
}

#NOFLG
function NOFLG {
    $msgTitle = "ERROR"
    $msgBody = "FLG File is Nothing in [C:\\cms5\\DB] !!"    
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END        
}

function CP_ERR {
    $msgTitle = "File Checking Error!!"
    $msgBody = "Not found W or G file of RCP`n`nPlease check in FD"   
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END    
}


#INPUT_ERROR
function INPUT_ERROR {
    $msgTitle = "ERROR"
    $msgBody = "SLC Over Error!!"   
    ShowMessageError $msgTitle $msgBody        
    $global:COMP_ERR = 1 
    THE_END
}

#NOT_READY
function NOT_READY {

    $msgTitle = "FD Drive Check Error!!"
    $msgBody = "Please insert froppy disk in $FD_DRV"   
    ShowMessageError $msgTitle $msgBody         
    CHKDRV    
}

#DISKERR
function DISKERR {
    $msgTitle = "Disk checking Error!!"
    $msgBody = "Disk read error!!"  
    ShowMessageError $msgTitle $msgBody     
    $global:COMP_ERR = 1
    THE_END
}

#########################################END_SYSTEM_MESSAGE###################################################

#########################################START_FUNCTION#######################################################
#Set location to work
function StartLocation {
    Set-Location $CURRENT_LOCATION
}

<#:/／￣￣ ＦＤドライブチェック ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
:/＜処理説明＞
:/　ＦＤドライブ内にＦＤが存在しているか
:/　ディスクに異常がないかチェックを行う。
#>
function CHKDRV { 
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        GETFN_NEW 
        return
    } 

    #Read driver info
    Get-PSDrive $FD_DRV

    if ($LASTEXITCODE -eq 10) {
        NOT_READY
    }
    elseif ($LASTEXITCODE -eq 1) {
        DISKERR
    }  
    
    GETFN_NEW
}

<#
GETFN_NEW
ＦＤよりＦＮファイル取得
ＦＤからＦＮファイルを取得し、選曲番号を生成する
作業曲に関係するLZHデータの保存場所名を生成する
#>
function GETFN_NEW {  
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {   
        ##Ver3.14 FN取得先を変更（キーボード入力＋FLGファイル有無）
        if ($ACOMP_SWITCH -ne 1) {
            $global:file_name = [string] (InputSongNumber "Please input song number Ex: 1234")

            if (([String]::IsNullOrWhiteSpace($global:file_name)) -eq $true) {
                $global:COMP_ERR = 1
                THE_END
                return
            }
        }

        #選曲番号を5 or 6桁に変換
        if ($global:file_name.Length -eq 1) {
            $global:SLC = "0000$global:file_name"
            $global:PSLC = "00000$global:file_name"
        }
        elseif ($global:file_name.Length -eq 2) {
            $global:SLC = "000$global:file_name"
            $global:PSLC = "0000$global:file_name"
        }
        elseif ($global:file_name.Length -eq 3) {
            $global:SLC = "00$global:file_name"
            $global:PSLC = "000$global:file_name"
        }
        elseif ($global:file_name.Length -eq 4) {
            $global:SLC = "0$global:file_name"
            $global:PSLC = "00$global:file_name"
        }
        elseif ($global:file_name.Length -eq 5) {
            $global:SLC = $global:file_name
            $global:PSLC = "0$global:file_name"
        }
        elseif ($global:file_name.Length -eq 6) {
            $global:SLC = $global:file_name
            $global:PSLC = $global:file_name
        }
        else {
            INPUT_ERROR
            return
        }    

        #国内制作ソフトハウス
        if ($global:METHOD_TYPE -eq 2) {
            $global:PSLC = ""
        }  

        GET_BK_DIR                      
    } 
    
    FLGDB_FILE
}

#GET_BK_DIR
function GET_BK_DIR {    
    [string] $bk_dir_tmp = GetFolderNumber $global:SLC
    
    if ($bk_dir_tmp.Length -eq 3) {
        $global:BK_DIR = "00$bk_dir_tmp"
    }
    elseif ($bk_dir_tmp.Length -eq 4) {
        $global:BK_DIR = "0$bk_dir_tmp"
    }
    elseif ($bk_dir_tmp.Length -eq 5) {
        $global:BK_DIR = $bk_dir_tmp
    }
    #Ver3.09追加　6桁対応
    elseif ($bk_dir_tmp.Length -eq 6) {
        #6桁の場合は選曲番号ディレクトリが１階層増える
        #Ver3.21修正(slc→bk_dir)
        $re = $bk_dir_tmp.Substring(0, 1)

        switch ($re) {
            1 {
                $re = "_10\"
                Break
            }
            2 {
                $re = "_20\"
                Break
            }  
            3 {
                $re = "_30\"
                Break
            }
            4 {
                $re = "_40\"
                Break
            }
            5 {
                $re = "_50\"
                Break
            }
            6 {
                $re = "_60\"
                Break
            }
            7 {
                $re = "_70\"
                Break
            }
            8 {
                $re = "_80\"
                Break
            }
            9 {
                $re = "_90\"
                Break
            }
            Default {
                Break
            }
        }

        $global:BK_DIR = "$re$bk_dir_tmp"
    }
    else {
        $msgInfo = "Back Up Directory Name Error!!"
        ShowMessageError "Error" $msgInfo
        $global:COMP_ERR = 1
        THE_END
    }

    #バックアップディレクトリが無い場合、作成する
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR")) -ne $true) {
        MakeDirectory ("$global:BK_DRV\$re")
        MakeDirectory ("$global:BK_DRV\$global:BK_DIR")
    }     
}

<#
:/／￣￣ FLG･DBファイル取得 ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
:/＜処理説明＞
:/　DBファイル、FLGファイルの有無を確認し、取得する
:/　FLG/DBNファイルが無い場合は処理を中断する
#>
function FLGDB_FILE {    
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {      
        #Ver3.14 入力した番号に0がついてる場合の対応
        $global:file_name = $global:file_name 

        #Ver3.14社外ネットワーク環境のみ
        #社外ネットワークはSドライブにFLGファイルが存在する
        if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
            #ネットワークから共通のフラグ置き場へコピー
            #FLG
            if ((CheckExistPath("$global:FLG_DIR\$global:file_name.FLG")) -eq $true) {

                COPY ("$global:FLG_DIR\$global:file_name.FLG") ("$DB_DIR\$global:file_name.FLG")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "FLG"
                    CP_ERR
                }

                #DBN
                #FLGフォルダからDBフォルダへコピー
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DBN")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DBN") ($DB_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DBN"
                        CP_ERR
                    }
                }
                #バックフォルダからDBフォルダへコピー
                elseif ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\7WK$global:SLC.LZH")) -eq $true) {
                    COPY ("$global:BK_DRV\$global:BK_DIR\7WK$global:SLC.LZH") ($TMP_DIR + "\7WK$global:SLC.LZH")

                    $param = "E $TMP_DIR\7WK$global:SLC", "$TMP_DIR\ *.*"
                    RUN_EXE $UNLHA_EXE_PATH $param
               
                    COPY ("$TMP_DIR\??$global:SLC" + "W.DBN") $DB_DIR

                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DBN"
                        CP_ERR
                    }
                }
                #バックフォルダからDBフォルダへコピー
                elseif ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\5WK$global:SLC.LZH")) -eq $true) {
                    COPY ("$global:BK_DRV\$global:BK_DIR\5WK$global:SLC.LZH") ($TMP_DIR + "\5WK$global:SLC.LZH")

                    $param = "E $TMP_DIR\5WK$global:SLC", "$TMP_DIR\ *.*"
                    RUN_EXE $UNLHA_EXE_PATH $param
                
                    COPY ("$TMP_DIR\??$global:SLC" + "W.DBN") $DB_DIR

                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DBN"
                        CP_ERR
                    }
                }
                #DBNが存在しない
                else {
                    NODBN_NETWORK
                }

                #:/DBN以外の映像系データはあれば持ってくる。
                #:/DDA
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DDA")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DDA") ($DB_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DDA"
                        CP_ERR
                    }
                }

                #:/DDB
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DDB")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DDB") ($DB_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DDB"
                        CP_ERR
                    }
                }

                #:/Ver3.23 DDC追加
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DDC")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DDC") ($DB_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DDC"
                        CP_ERR
                    }
                }

                #:/er3.40 DDE追加               
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DDE")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DDE") ($DB_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DDE"
                        CP_ERR
                    }
                }

                #:/Ver3.40 DBM追加               
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DBM")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DBM") ($DB_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DBM"
                        CP_ERR
                    }
                }

                ##########START_INS_2_検査ソフトハウス###################
                #Ver4.67追加
                #BMN               
                if ((CheckExistPath("$global:FLG_DIR\$global:SLC.BMN")) -eq $true) {

                    COPY("$global:FLG_DIR\$global:SLC.BMN") ("$DB_DIR")

                    if ($? -eq $false) {
                        $global:ERR_FNAME = "BMN"
                        CP_ERR
                    }
                }
                                 
                ##########END_INS_2_検査ソフトハウス###################

                #:/Ver3.15 CSV,PDF,TXT追加
                # CSV               
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.CSV")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.CSV") ($TMP_DIR)

                    $fileTmpPath = "$TMP_DIR\??$global:SLC" + "W.CSV"
                    $fileTmpPath = "$TMP_DIR\a.txt"
                    SaveFileInfo $fileTmpPath $fileTmpPath

                    #Get data in line index
                    $sh_tmp = GetContentByIndex $fileTmpPath 0                                     
                    DeleteFileOrFolder $fileTmpPath

                    $global:sh = $sh_tmp.Substring(0, 2);   

                    if ($? -eq $false) {
                        $global:ERR_FNAME = "CSV"
                        CP_ERR
                    }
                }

                #:/PDF
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.PDF")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.PDF") ($TMP_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "PDF"
                        CP_ERR
                    }
                }

                #:/TXT
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.TXT")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.TXT") ($TMP_DIR)
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "TXT"
                        CP_ERR
                    }
                }
                
                ##########################START_INS_3_国内制作ソフトハウス############################################
                #DLF(Ver 5.02)
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DLF")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DLF") ("$DB_DIR")
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DLF"
                        CP_ERR
                    }
                }

                #DLG(Ver 5.02)
                if ((CheckExistPath("$global:FLG_DIR\??$global:SLC" + "W.DLG")) -eq $true) {
                    COPY ("$global:FLG_DIR\??$global:SLC" + "W.DLG") ("$DB_DIR")
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "DLG"
                        CP_ERR
                    }
                }                              
                ##########################END_INS_3_国内制作ソフトハウス############################################

                #Ver3.35 Key情報、演奏時間が記入されたCSVファイルを追加
                if ((CheckExistPath("$global:RCP_DIR\??$global:SLC.CSV")) -eq $true) {
                    COPY ("$global:RCP_DIR\??$global:SLC.CSV") ("$TMP_DIR\get.csv")
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "CSV_"
                        CP_ERR
                    }
                }
            }
            #FLGファイルがローカルに存在すれば、INSは可能とする
            elseif ((CheckExistPath("$DB_DIR\$global:file_name" + ".FLG")) -eq $true) {

                #Ver3.35 Key情報、演奏時間が記入されたCSVファイルを追加
                if ((CheckExistPath("$RCP_DIR\??$global:SLC" + ".CSV")) -eq $true) {
                    COPY ("$RCP_DIR\??$global:SLC" + ".CSV") ("$TMP_DIR\get.csv")
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "CSV_"
                        CP_ERR
                    }
                }
            }
            #FLGファイルがネットワークにない場合、INSは不可とする
            else {
                NOFLG_NETWORK
            }

            #(Ver3.28追加)音多Flag
            if ((CheckExistPath("$global:FLG_DIR\$global:file_name.ONT")) -eq $true) {
                COPY ("$global:FLG_DIR\$global:file_name.ONT") ("$DB_DIR\$global:file_name.ONT")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "ONT"
                    CP_ERR
                }
            }
        }
      
        #Write-Host "aaaaa: $DB_DIR\$global:file_name.FLG"
        # return
        #FLGファイル有無確認
        if ((CheckExistPath("$DB_DIR\$global:file_name.FLG")) -eq $true) { 
            #DBNファイル名を出力(発注区分取得の為)
            if ((CheckExistPath("$DB_DIR\??$global:SLC" + "W.DBN")) -eq $true) {
                $soureFile = "$DB_DIR\??$global:SLC" + "W.DBN"
                $desFile = "$TMP_DIR\$TMP_NAME"
                SaveFileInfo $soureFile $desFile 
            }
            #バックフォルダからDBNを取得
            elseif ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\7WK$global:SLC.LZH")) -eq $true) {
                COPY ("$global:BK_DRV\$global:BK_DIR\7WK$global:SLC.LZH") ("$TMP_DIR\7WK$global:SLC.LZH")
                $param = "E $TMP_DIR\7WK$global:SLC $TMP_DIR\ *.DBN"
                RUN_EXE $UNLHA_EXE_PATH $param               
             
                COPY ("$TMP_DIR\??$global:SLC" + "W.DBN") ("$DB_DIR")

                #save file info
                $soureFile = "$DB_DIR\??$global:SLC" + "W.DBN"
                $desFile = "$TMP_DIR\$TMP_NAME"

                SaveFileInfo $soureFile $desFile 
               
                DeleteFileOrFolder("$TMP_DIR\7WK$global:SLC.LZH")
            }
            elseif ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\5WK$global:SLC.LZH")) -eq $true) {

                COPY ("$global:BK_DRV\$global:BK_DIR\5WK$global:SLC.LZH") ("$TMP_DIR\5WK$global:SLC.LZH")

                $param = "E $TMP_DIR\5WK$global:SLC $TMP_DIR\ *.DBN"
                RUN_EXE $UNLHA_EXE_PATH $param

                #Ver3.30追加
                if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\AWK$global:SLC.LZH")) -eq $true) {
                    COPY ("$global:BK_DRV\$global:BK_DIR\AWK$global:SLC.LZH") ("$TMP_DIR\AWK$global:SLC.LZH")

                    $param = "E -m1 $TMP_DIR\AWK$global:SLC $TMP_DIR\ *.DBN"
                    RUN_EXE $UNLHA_EXE_PATH $param
                }

                if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\CWK$global:SLC.LZH")) -eq $true) {
                    COPY ("$global:BK_DRV\$global:BK_DIR\CWK$global:SLC.LZH") ("$TMP_DIR\CWK$global:SLC.LZH")

                    $param = "E -m1 $TMP_DIR\CWK$global:SLC $TMP_DIR\ *.DBN"
                    RUN_EXE $UNLHA_EXE_PATH $param
                }
               
                COPY("$TMP_DIR\??$global:SLC" + "W.DBN") ("$DB_DIR")

                if ($? -eq $false) {
                    $global:ERR_FNAME = "DBN" 
                    CP_ERR
                }

                #save file info              
                $soureFile = "$DB_DIR\??$global:SLC" + "W.DBN"
                $desFile = "$TMP_DIR\$TMP_NAME"

                SaveFileInfo $soureFile $desFile 

                DeleteFileOrFolder("$TMP_DIR\5WK$global:SLC.LZH")
                DeleteFileOrFolder("$TMP_DIR\AWK$global:SLC.LZH")
                DeleteFileOrFolder("$TMP_DIR\CWK$global:SLC.LZH")

                DeleteFileOrFolder("$TMP_DIR\??$global:SLC?.???")   
            }
            #DBNが存在しない
            else {
                NODBN
            }
        }
        else {          
            NOFLG            
        }

        OUT_RCP_NAME
        return
    }

    #社外FDのみ
    if ((CheckExistPath("$FD_DRV\*.FN")) -ne $true) {
        NO_FN_FILE
    }

    #Write-Host "FLGDB_FILE-OUT_RCP_NAME"
    SaveFileInfo ("$FD_DRV\*.FN") ("$TMP_DIR\$TMP_NAME")

    OUT_RCP_NAME
}

#OUT_RCP_NAME
function OUT_RCP_NAME { 
    $RCP_NAME = [string] (GetContentByIndex ("$TMP_DIR\$TMP_NAME") 0)
   
    if ([string]::IsNullOrWhiteSpace($RCP_NAME)) {
        NO_TMP_FILE_ERR
    }

    #5桁の場合(FNより)
    if ($global:METHOD_TYPE -eq 1 -and $RCP_NAME.Length -eq 10) {
        $global:SLC = $RCP_NAME.Substring(2, 5)
        $global:FN_HEAD = $RCP_NAME.Substring(0, 2)        
        $global:FN = $RCP_NAME.Substring(0, 7)
    }
    #5桁の場合(DBNより)
    elseif ($RCP_NAME.Length -eq 12) {
        $global:SLC = $RCP_NAME.Substring(2, 5)
        $global:FN_HEAD = $RCP_NAME.Substring(0, 2)        
        $global:FN = $RCP_NAME.Substring(0, 7)
    }
    #6桁の場合
    else {
        $global:SLC = $RCP_NAME.Substring(2, 6)
        $global:FN_HEAD = $RCP_NAME.Substring(0, 2)        
        $global:FN = $RCP_NAME.Substring(0, 8)
    }

    #:/Ver3.14追加 Auto時はSELECT.TMPの削除を行なわない
    if ($ACOMP_SWITCH -ne 1) {       
        DeleteFileOrFolder("$TMP_DIR\$TMP_NAME")
    }
   
    GET_BK_DIR   

    #Ver3.14 社外FDのみ実行
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {       
        RCPLHA  
        return
    }

    $global:DRV = "$FD_DRV\88"
    $global:RN = "R36"
    GET_RCPREV

    $global:R88_N = $global:RN
    $global:R88_REV = $global:REV

    if ((CheckExistPath("$FD_DRV\55\*.*")) -eq $true) {
        $global:DRV = "$FD_DRV\55"
        $global:RN = "R36"
        GET_RCPREV

        $global:R55_N = $global:RN
        $global:R55_REV = $global:REV
    }

    if ((CheckExistPath("$FD_DRV\AA\*.R36")) -eq $true) {
        $global:DRV = "$FD_DRV\AA"
        $global:RN = "R36"
        GET_RCPREV

        $global:RAA_N = $global:RN
        $global:RAA_REV = $global:REV
    }
    if ((CheckExistPath("$FD_DRV\BB\*.R36")) -eq $true) {
        $global:DRV = "$FD_DRV\BB"
        $global:RN = "R36"
        GET_RCPREV

        $global:RBB_N = $global:RN
        $global:RBB_REV = $global:REV
    }
    #/Ver3.23追加
    if ((CheckExistPath("$FD_DRV\CC\*.R36")) -eq $true) {
        $global:DRV = "$FD_DRV\CC"
        $global:RN = "R36"
        GET_RCPREV

        $global:RCC_N = $global:RN
        $global:RCC_REV = $global:REV
    }

    CPRCP_NEW
}

##RCPLHA
function RCPLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\RCP$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\RCP$global:SLC.LZH") ("$TMP_DIR\RCP$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "RCPLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\RCP$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }

    8RCLHA
}

#8RCLHA
function 8RCLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\8RC$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\8RC$global:SLC.LZH") ("$TMP_DIR\8RC$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "8RCLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\8RC$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    PLYLHA
}

#PLYLHA
function PLYLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLY$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLY$global:SLC.LZH") ("$TMP_DIR\PLY$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "PLYLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\PLY$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    } 
    
    8PLLHA
}

#8PLLHA
function 8PLLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\8PL$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\8PL$global:SLC.LZH") ("$TMP_DIR\8PL$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "8PLLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\8PL$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    } 
    
    HDTLHA
}

#HDTLHA
function HDTLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\HDT$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\HDT$global:SLC.LZH") ("$TMP_DIR\HDT$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "HDTLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\HDT$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }  
    
    PDFLHA
}

#PDFLHA
function PDFLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PDF$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PDF$global:SLC.LZH") ("$TMP_DIR\PDF$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "PDFLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\PDF$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    5WKLHA
}

#5WKLHA
function 5WKLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\5WK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\5WK$global:SLC.LZH") ("$TMP_DIR\5WK$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "5WKLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\5WK$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    FMDLHA
}


#FMDLHA
function FMDLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\FMD$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\FMD$global:SLC.LZH") ("$TMP_DIR\FMD$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "FMDLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\FMD$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    CPDLHA
}


#FMDLHA
function CPDLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\CPD$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\CPD$global:SLC.LZH") ("$TMP_DIR\CPD$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "CPDLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\CPD$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }    

    TARLHA  
}

#TARLHA
function TARLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\TAR$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\TAR$global:SLC.LZH") ("$TMP_DIR\TAR$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "TARLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\TAR$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }
    
    DIGLHA
}


#DIGLHA
function DIGLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\DIG$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\DIG$global:SLC.LZH") ("$TMP_DIR\DIG$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "DIGLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\DIG$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }  
    
    PLALHA
}


#:/ ----- Ver3.04(JS70対応) 追加 ----------------------
#DIGLHA
function PLALHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLA$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLA$global:SLC.LZH") ("$TMP_DIR\PLA$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "PLALZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\PLA$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   

    AWKLHA   
}

#AWKLHA
function AWKLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\AWK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\AWK$global:SLC.LZH") ("$TMP_DIR\AWK$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "AWKLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\AWK$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }       
  
    WAVLHA 2
}

#:/ソフトハウス環境のみコピーしてくる
#WAVLHA
function WAVLHA { 
    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\WAV$global:SLC.LZH")) -eq $true) {
            COPY ("$global:BK_DRV\$global:BK_DIR\WAV$global:SLC.LZH") ("$TMP_DIR\WAV$global:SLC.LZH")

            if ($? -eq $false) {
                $global:ERR_FNAME = "WAVLZH"
                CP_ERR
                return
            }
            $param = "E $TMP_DIR\WAV$global:SLC", "$TMP_DIR\ *.*"
            RUN_EXE $UNLHA_EXE_PATH $param    
        }     
    } 

    PLBLHA
}

#:/ ----- Ver3.06(BB対応) 追加 ----------------------
#PLBLHA
function PLBLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLB$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLB$global:SLC.LZH") ("$TMP_DIR\PLB$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "PLBLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\PLB$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    PLCLHA
}

#:/ ----- Ver3.23(JS-W1対応) 追加 ----------------------
#PLBLHA
function PLCLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLC$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLC$global:SLC.LZH") ("$TMP_DIR\PLC$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "PLCLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\PLC$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    CWKLHA
}

#PLBLHA
function CWKLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\CWK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\CWK$global:SLC.LZH") ("$TMP_DIR\CWK$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "CWKLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\CWK$global:SLC" , "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    9PLHA
}

#:/ ----- Ver3.24(色替りPV対応) 追加 ----------------------
#9PLHA
function 9PLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\9PL$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\9PL$global:SLC.LZH") ("$TMP_DIR\9PL$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "9PLLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\9PL$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    9WKLHA
}

#:/ ----- Ver3.33(XJ-J1対応) 追加 ----------------------
#9WKLHA
function 9WKLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\9WK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\9WK$global:SLC.LZH") ("$TMP_DIR\9WK$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "9WKLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\9WK$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    PLDLHA
}

#:/ ----- Ver3.33(XJ-J1対応) 追加 ----------------------
#PLDLHA
function PLDLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLD$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLD$global:SLC.LZH") ("$TMP_DIR\PLD$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "PLDLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\PLD$global:SLC" , "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    CMDLHA
}

#:/ ----- Ver3.34(共通SMF対応) 追加 ----------------------
#CMDLHA
function CMDLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\CMD$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\CMD$global:SLC.LZH") ("$TMP_DIR\CMD$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "CMDLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\CMD$global:SLC" , "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    EWKLHA
}

#:/ ----- Ver3.40(JS-WX対応) 追加 ----------------------
#EWKLHA
function EWKLHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\EWK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\EWK$global:SLC.LZH") ("$TMP_DIR\EWK$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "EWKLZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\EWK$global:SLC" , "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   
    
    PLELHA
}


#:/ ----- Ver3.40(JS-WX対応) 追加 ----------------------
#PLELHA
function PLELHA {
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLE$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLE$global:SLC.LZH") ("$TMP_DIR\PLE$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "PLELZH"
            CP_ERR
            return
        }
        $param = "E $TMP_DIR\PLE$global:SLC" , "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }   

    AACLHA 
}


#:/ ----- Ver3.40(JS-WX対応) 追加 ----------------------
#AACLHA
function AACLHA {   

    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\AAC$global:SLC.LZH")) -eq $true) {
            COPY ("$global:BK_DRV\$global:BK_DIR\AAC$global:SLC.LZH") ("$TMP_DIR\AAC$global:SLC.LZH")

            if ($? -eq $false) {
                $global:ERR_FNAME = "AACLZH"
                CP_ERR
                return
            }
            $param = "E $TMP_DIR\AAC$global:SLC", "$TMP_DIR\ *.*"
            RUN_EXE $UNLHA_EXE_PATH $param    
        }   
    }
    
    SPEAAC
}

#SPEAAC
function SPEAAC {
    if ((CheckExistPath("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "GM0.AAC")) -eq $true) {
        COPY ("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "GM0.AAC") ("$TMP_DIR\$global:FN" + "GM0.AAC")

        if ($? -eq $false) {
            $global:ERR_FNAME = "AAC"
            CP_ERR
            return
        }       
    }  

    if ((CheckExistPath("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "ET0.AAC")) -eq $true) {
        COPY ("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "ET0.AAC") ("$TMP_DIR\$global:FN" + "ET0.AAC")

        if ($? -eq $false) {
            $global:ERR_FNAME = "AAC"
            CP_ERR
            return
        }       
    }  

    if ((CheckExistPath("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "GT0.AAC")) -eq $true) {
        COPY ("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "GT0.AAC") ("$TMP_DIR\$global:FN" + "GT0.AAC")

        if ($? -eq $false) {
            $global:ERR_FNAME = "AAC"
            CP_ERR
            return
        }       
    }  

    if ((CheckExistPath("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "DR0.AAC")) -eq $true) {
        COPY ("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "DR0.AAC") ("$TMP_DIR\$global:FN" + "DR0.AAC")

        if ($? -eq $false) {
            $global:ERR_FNAME = "AAC"
            CP_ERR
            return
        }       
    }  

    if ((CheckExistPath("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "ET1.AAC")) -eq $true) {
        COPY ("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "ET1.AAC") ("$TMP_DIR\$global:FN" + "ET1.AAC")

        if ($? -eq $false) {
            $global:ERR_FNAME = "ET1AAC"
            CP_ERR
            return
        }       
    }  

    if ((CheckExistPath("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "GM1.AAC")) -eq $true) {
        COPY ("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "GM1.AAC") ("$TMP_DIR\$global:FN" + "GM1.AAC")

        if ($? -eq $false) {
            $global:ERR_FNAME = "GM1AAC"
            CP_ERR
            return
        }       
    }  

    ##########START_INS_1_XING社内###################
    #Ver4.67追加
    #AAC  
    if ((CheckExistPath("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "ET2.AAC")) -eq $true) {
        COPY ("$global:AAC_DIR\$global:BK_DIR\WX$global:SLC" + "ET2.AAC") ("$TMP_DIR\$global:FN" + "ET2.AAC")
                
        if ($? -eq $false) {
            $global:ERR_FNAME = "ET2AAC"
            CP_ERR
            return
        }       
    }                         
    ##########END_INS_1_XING社内###################

    7WKLHA
}

#:/ ----- 以下、Ver4.00(新AT対応)  -------------------------
#7WKLHA
function 7WKLHA {   
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\7WK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\7WK$global:SLC.LZH") ("$TMP_DIR\7WK$global:SLC.LZH")

        Write-Host "COPY: $path"
        if ($? -eq $false) {

            Write-Host "ERROR: $path"
            $global:ERR_FNAME = "7WKLZH"
            CP_ERR
            return
        } 
        
        $param = " E $TMP_DIR\7WK$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }  

    UGALHA
}

#:/ ----- 以下、Ver4.02追加 --------------------------------------
#UGALHA
function UGALHA { 

    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\UGA$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\UGA$global:SLC.LZH") ("$TMP_DIR\UGA$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "UGALZH"
            CP_ERR
            return
        } 
        
        $param = "E $TMP_DIR\UGA$global:SLC", "$TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }  

    MSYLHA 
}


#:/ ----- 以下、Ver4.03追加 --------------------------------------
#MSYLHA
function MSYLHA {  
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\MSY$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\MSY$global:SLC.LZH") ("$TMP_DIR\MSY$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "MSYLZH"
            CP_ERR
            return
        } 
        
        $param = "E $TMP_DIR\MSY$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }  

    #test
    FWKLHA 
}


#:/ ----- 以下、Ver4.20追加 --------------------------------------
#FWKLHA
function FWKLHA {
  
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\FWK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\FWK$global:SLC.LZH") ("$TMP_DIR\FWK$global:SLC.LZH")

        if ($? -eq $false) {
            $global:ERR_FNAME = "FWKLZH"
            CP_ERR
            return
        } 
        
        $param = "E $TMP_DIR\FWK$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
      
        #\\JS-server\Kanri のR01が新しければそちらを優先する
        if ($global:METHOD_TYPE -eq 0) {
            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R01.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R01.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R01mid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R01_B.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R01_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R01mid"
                    CP_ERR
                    return
                } 
            }

            #\\JS-server\Kanri のRFBが新しければそちらを優先する
            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFB.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFB.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RFBmid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFB_B.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFB_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RFBmid"
                    CP_ERR
                    return
                } 
            }
            #\\JS-server\Kanri のRFNが新しければそちらを優先する
            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFN.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFN.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RFNmid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFN_B.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RFN_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RFNmid"
                    CP_ERR
                    return
                } 
            }
        }
    }
    else {
        #/KENDATAを見るのは、社内だけ
        if ($global:METHOD_TYPE -eq 0) {
            if ((CheckExistPath("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_R01.mid")) -eq $true) {
                COPY ("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_R01.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R01mid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_R01_B.mid")) -eq $true) {
                COPY ("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_R01_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R01mid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_RFB.mid")) -eq $true) {
                COPY ("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_RFB.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RFBmid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_RFB_B.mid")) -eq $true) {
                COPY ("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_RFB_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RFBmid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_RFN.mid")) -eq $true) {
                COPY ("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W_RFN.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RFNmid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W.MPT")) -eq $true) {
                COPY ("$global:KEN_DIR\MID\$global:BK_DIR\$global:FN" + "W.MPT") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "MPTmid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\SEA\$global:BK_DIR\$global:FN" + "sea.mid")) -eq $true) {
                COPY ("$global:KEN_DIR\SEA\$global:BK_DIR\$global:FN" + "sea.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "sea"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\AAC\$global:BK_DIR\WX$global:SLC" + "ET1.aac")) -eq $true) {
                COPY ("$global:KEN_DIR\AAC\$global:BK_DIR\WX$global:SLC" + "ET1.aac") ("$TMP_DIR\$global:FN" + "ET1.aac")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "ET1AAC"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\AAC\$global:BK_DIR\WX$global:SLC" + "GM1.aac")) -eq $true) {
                COPY ("$global:KEN_DIR\AAC\$global:BK_DIR\WX$global:SLC" + "GM1.aac") ("$TMP_DIR\$global:FN" + "GM1.aac")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "GM1AAC"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KEN_DIR\DBTXT\$global:BK_DIR\$global:FN" + "_KenTTL.txt")) -eq $true) {
                COPY ("$global:KEN_DIR\DBTXT\$global:BK_DIR\$global:FN" + "_KenTTL.txt") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "KENTTL"
                    CP_ERR
                    return
                } 
            }
        }
    }

    PLFLHA 
}

#PLFLHA
function PLFLHA {   

    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLF$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLF$global:SLC.LZH") ("$TMP_DIR\PLF$global:SLC.LZH")
        if ($? -eq $false) {
            $global:ERR_FNAME = "PLFLZH"
            CP_ERR
            return
        } 

        $param = "E $TMP_DIR\PLF$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param    
    }

    GWKLHA 
}

#:/ ----- 以下、Ver4.61追加 ZEUS / ATHENA 対応 --------------------------------------
#GWKLHA
function GWKLHA {   
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\GWK$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\GWK$global:SLC.LZH") ("$TMP_DIR\GWK$global:SLC.LZH")
        if ($? -eq $false) {
            $global:ERR_FNAME = "GWKLZH"
            CP_ERR
            return
        } 

        $param = "E $TMP_DIR\GWK$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param  
                
        if ($global:METHOD_TYPE -eq 0) {
            #Zeus_Dataの方が新しければそちらを優先する
            #MPT
            if ((CheckExistPath("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W.MPT")) -eq $true) {
                COPY ("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W.MPT") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "MPTmid"
                    CP_ERR
                    return
                } 
            }

            #\\JS-server\Kanri のR02が新しければそちらを優先する
            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R02.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R02.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R02mid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R02_B.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_R02_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R02mid"
                    CP_ERR
                    return
                } 
            }

            #\\JS-server\Kanri のRB2が新しければそちらを優先する
            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RB2.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RB2.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RB2mid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RB2_B.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RB2_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RB2mid"
                    CP_ERR
                    return
                } 
            }

            #\\JS-server\Kanri のRN2が新しければそちらを優先する
            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RN2.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RN2.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RN2mid"
                    CP_ERR
                    return
                } 
            }

            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RN2_B.mid")) -eq $true) {
                COPY ("$global:KANRI_D\MP2WAVE\Data\$global:SLC\$global:FN" + "W_RN2_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RN2mid"
                    CP_ERR
                    return
                } 
            }
        }
    }
    else {
        #:/ZEUS_DATAを見るのは、社内だけ
        if ($global:METHOD_TYPE -eq 0) {

            #:/MPT
            if ((CheckExistPath("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W.MPT")) -eq $true) {
                COPY ("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W.MPT") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "MPTmid"
                    CP_ERR
                    return
                } 
            }

            #R02
            if ((CheckExistPath("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_R02.mid")) -eq $true) {
                COPY ("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_R02.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R02mid"
                    CP_ERR
                    return
                } 
            }

            #RB2
            if ((CheckExistPath("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_RB2.mid")) -eq $true) {
                COPY ("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_RB2.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RB2mid"
                    CP_ERR
                    return
                } 
            }

            #MPT_B
            if ((CheckExistPath("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_B.MPT")) -eq $true) {
                COPY ("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_B.MPT") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "MPTBmid"
                    CP_ERR
                    return
                } 
            }

            #R02_B
            if ((CheckExistPath("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_R02_B.mid")) -eq $true) {
                COPY ("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_R02_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R02Bmid"
                    CP_ERR
                    return
                } 
            }

            #RB2_B
            if ((CheckExistPath("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_RB2_B.mid")) -eq $true) {
                COPY ("$global:ZEUS_DIR\MID_MPT\$global:BK_DIR\$global:FN" + "W_RB2_B.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RB2Bmid"
                    CP_ERR
                    return
                } 
            }

            #AVE
            if ((CheckExistPath("$global:ZEUS_DIR\AVE\AVE\CG0$global:PSLC.AVE")) -eq $true) {
                COPY ("$global:ZEUS_DIR\AVE\AVE\CG0$global:PSLC.AVE") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "AVE"
                    CP_ERR
                    return
                } 
            }

            #VOL
            if ((CheckExistPath("$global:ZEUS_DIR\VOL\$global:BK_DIR\C$global:PSLC.vol")) -eq $true) {
                COPY ("$global:ZEUS_DIR\VOL\$global:BK_DIR\C$global:PSLC.vol") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "VOL"
                    CP_ERR
                    return
                } 
            }
        }
    }

    PLGLHA 
}

#PLGLHA
function PLGLHA {  
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\PLG$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\PLG$global:SLC.LZH") ("$TMP_DIR\PLG$global:SLC.LZH")
        if ($? -eq $false) {
            $global:ERR_FNAME = "PLGLZH"
            CP_ERR
            return
        } 

        $param = "E $TMP_DIR\PLG$global:SLC $TMP_DIR\ *.*"
        RUN_EXE $UNLHA_EXE_PATH $param  
    }

    SSYLHA 
}

#:/ ----- 以下、Ver4.23追加 --------------------------------------
#:/ ----- SSY  -------------------------
function SSYLHA {  
   
    if ((CheckExistPath("$global:BK_DRV\$global:BK_DIR\SSY$global:SLC.LZH")) -eq $true) {
        COPY ("$global:BK_DRV\$global:BK_DIR\SSY$global:SLC.LZH") ("$TMP_DIR\SSY$global:SLC.LZH")
        if ($? -eq $false) {
            $global:ERR_FNAME = "SSYLZH"
            CP_ERR
            return
        } 

        $param = "E $TMP_DIR\SSY$global:SLC $TMP_DIR\ *.XLM"
        RUN_EXE $UNLHA_EXE_PATH $param  

        $param = "E $TMP_DIR\SSY$global:SLC $TMP_DIR\ *.EFF"
        RUN_EXE $UNLHA_EXE_PATH $param  
    }   

    DELLZH 
}

#:/ ------------------------------------ Ver4.23追加追加終了 -----
#DELLZH
function DELLZH {  

    if ((CheckExistPath("$TMP_DIR\???$global:SLC.LZH")) -eq $true) {
        DeleteFileOrFolder("$TMP_DIR\???$global:SLC.LZH")
    }

    #:/ Ver1.07で修正
    #:/FD環境では行わない
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        if ((CheckExistPath("$global:KANRI_D\TTL\??$global:SLC" + "W.TTB")) -eq $true) {
            COP("$global:KANRI_D\TTL\??$global:SLC" + "W.TTB") ("$TMP_DIR\$global:FN" + "W.TTB")

            if ($? -eq $false) {
                $global:ERR_FNAME = "TTB"
                CP_ERR
                return
            } 
        }
        elseif ((CheckExistPath("$global:KANRI_D\TTL\??$global:SLC" + "W.TTA")) -eq $true) {
            COP("$global:KANRI_D\TTL\??$global:SLC" + "W.TTA") ("$TMP_DIR\$global:FN" + "W.TTA")

            if ($? -eq $false) {
                $global:ERR_FNAME = "TTA"
                CP_ERR
                return
            } 
        }
        elseif ((CheckExistPath("$global:KANRI_D\TTL\??$global:SLC" + "W.TTL")) -eq $true) {
            COP("$global:KANRI_D\TTL\??$global:SLC" + "W.TTL") ("$TMP_DIR\$global:FN" + "W.TTL")

            if ($? -eq $false) {
                $global:ERR_FNAME = "TTL"
                CP_ERR
                return
            } 
        }

        if ((CheckExistPath("$global:KANRI_D\TTL\$global:SLC.TIM")) -eq $true) {
            COP("$global:KANRI_D\TTL\$global:SLC.TIM") ("$TMP_DIR\$global:SLC.TIM")

            if ($? -eq $false) {
                $global:ERR_FNAME = "TIM"
                CP_ERR
                return
            } 
        }

        if ((CheckExistPath("$global:KANRI_D\TTL\??$global:SLC" + "W.MMR")) -eq $true) {
            COP("$global:KANRI_D\TTL\??$global:SLC" + "W.MMR") ("$TMP_DIR\$global:FN" + "W.MMR")

            if ($? -eq $false) {
                $global:ERR_FNAME = "MMR"
                CP_ERR
                return
            } 
        }
        
        if ((CheckExistPath("$global:KANRI_D\TTL\??$global:SLC" + "W.BMP")) -eq $true) {
            COP("$global:KANRI_D\TTL\??$global:SLC" + "W.BMP") ("$TMP_DIR\$global:FN" + "W.BMP")

            if ($? -eq $false) {
                $global:ERR_FNAME = "BMP"
                CP_ERR
                return
            } 
        }       
    
        if ((CheckExistPath("$global:KANRI_D\TTL\??$global:SLC" + "W.PTN")) -eq $true) {
            COP("$global:KANRI_D\TTL\??$global:SLC" + "W.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")

            if ($? -eq $false) {
                $global:ERR_FNAME = "PTN"
                CP_ERR
                return
            } 
        }

        if ((CheckExistPath("$global:KANRI_D\TTL\??$global:SLC" + "W.TXT")) -eq $true) {
            COP("$global:KANRI_D\TTL\??$global:SLC" + "W.TXT") ("$TMP_DIR\$global:FN" + "W.TXT")

            if ($? -eq $false) {
                $global:ERR_FNAME = "TXT"
                CP_ERR
                return
            } 
        }

        #ＡＤＰ専用ファイル準備
        if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC" + "?.ADP")) -eq $true) {
            $global:EXTENSION = "ADP"            
        }
        elseif ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC" + "?.APC")) -eq $true) {
            $global:EXTENSION = "APC"    
        }
        else {
            CPYAPNEXT 
            return;
        }

        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.APC")
            }
        }

        CPYAPNEXT 
    }
}

#:/ ----- Ver3.05追加 ----------------------
#CPYAPNEXT
function CPYAPNEXT {    
   
    if ((CheckExistPath("$global:KANRI_D\ADP\??$global:SLC" + "?.ADP")) -eq $true) {
        $global:EXTENSION = "ADP"            
    }
    elseif ((CheckExistPath("$global:KANRI_D\ADP\??$global:SLC" + "?.APC")) -eq $true) {
        $global:EXTENSION = "APC"    
    }
    else {
        CPYACMs
        return;
    }

    $ARRAYS | foreach {
        if ((CheckExistPath("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION")) -eq $true) {
            COPY("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.APC")
        }
    }

    CPYACMs 
}

#CPYACMs
function CPYACMs {  
    #ACM
    if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC" + "?.ACM")) -eq $true) {
        $global:EXTENSION = "ACM"     
        
        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }

        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }
    }

    #ABM
    if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC" + "?.ABM")) -eq $true) {
        $global:EXTENSION = "ABM"     
        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }

        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }
    }

    #ACS
    if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC" + "?.ACS")) -eq $true) {
        $global:EXTENSION = "ACS"     
        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }

        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }
    }

    #ABS
    if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC" + "?.ABS")) -eq $true) {
        $global:EXTENSION = "ABS"     
        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:SLC$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }

        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION")) -eq $true) {
                COPY("$global:KANRI_D\ADP\$global:FN$_.$global:EXTENSION") ("$TMP_DIR\$global:FN$_.$global:EXTENSION")
            }
        }
    }

    CPYKDB 
}

#:/ ----- 以上、Ver3.05追加終了 ----------------------
#CPYKDB
function CPYKDB {   
    #DBN
    if ((CheckExistPath("$global:KANRI_D\DB\??$global:SLC" + "W.DBN")) -eq $true) {
        COPY("$global:KANRI_D\DB\??$global:SLC" + "W.DBN") ("$TMP_DIR\$global:FN" + "W.DBN")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DBN"
            CP_ERR
            return
        } 
    }
    #DDB
    if ((CheckExistPath("$global:KANRI_D\DDB\??$global:SLC" + "W.DDB")) -eq $true) {
        COPY("$global:KANRI_D\DDB\??$global:SLC" + "W.DDB") ("$TMP_DIR\$global:FN" + "W.DDB")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DDB"
            CP_ERR
            return
        } 
    }

    DDACOPY 
}

#DDACOPY
function DDACOPY {   

    #DDA
    if ((CheckExistPath("$global:KANRI_D\DDA\??$global:SLC*.DDA")) -eq $true) {
        COPY("$global:KANRI_D\DDA\??$global:SLC*.DDA") ("$TMP_DIR\$global:FN" + "W.DDA")       
    }  
    
    1DTCOPY 
}

#1DTCOPY
function 1DTCOPY { 
    #DDA
    if ((CheckExistPath("$global:KANRI_D\1DT\??$global:SLC*.1DT")) -eq $true) {
        COPY("$global:KANRI_D\1DT\??$global:SLC*.1DT") ("$TMP_DIR\$global:FN" + "W.1DT")       
    }  
    
    TDBCOPY 
}

#TDBCOPY
function TDBCOPY {    
    #TDB
    if ((CheckExistPath("$global:KANRI_D\TDB\??$global:SLC*.TDB")) -eq $true) {
        COPY("$global:KANRI_D\TDB\??$global:SLC*.TDB") ("$TMP_DIR\$global:FN" + "W.TDB")       
    }  

    if ((CheckExistPath("$global:KANRI_D\JPEG\??$global:SLC*.JPG")) -eq $true) {
        COPY("$global:KANRI_D\JPEG\??$global:SLC*.JPG") ("$TMP_DIR\$global:FN" + "W.JPG")       
    }
    elseif ((CheckExistPath("$global:KANRI_D\BMA2\??$global:SLC*.BMP")) -eq $true) {
        COPY("$global:KANRI_D\BMA2\??$global:SLC*.BMP") ("$TMP_DIR\$global:FN" + "W.BMA")       
    }
    elseif ((CheckExistPath("$global:KANRI_D\BMA\??$global:SLC*.BMA")) -eq $true) {
        COPY("$global:KANRI_D\BMA\??$global:SLC*.BMA") ("$TMP_DIR\$global:FN" + "W.BMA")       
    } 

    MP2COPY 
}

#:/Ver3.29
#MP2COPY
function MP2COPY {   

    $ARRAYS | foreach {
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC$_.MP2")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC$_.MP2") ("$TMP_DIR\$global:FN$_.MP2")
        }
    }

    APCCOPY 
}

#:/Ver3.26追加
#:_APCCOPY
function APCCOPY {  
    $ARRAYS | foreach {
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC$_.APC")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC$_.APC") ("$TMP_DIR\$global:FN$_.APC")
        }
    }

    MIDCOPY 
}

#MIDCOPY
function MIDCOPY {     
    if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.MPT")) -eq $true) {
        COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.MPT") ("$TMP_DIR\$global:FN" + "W.MPT")       
    }

    RAACOPY 
}

#RAACOPY
function RAACOPY {   
    if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.RAA")) -eq $true) {
        COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.RAA") ("$TMP_DIR\$global:FN" + "W.RAA")       
    }
    
    CWPCOPY
}

#CWPCOPY
function CWPCOPY {  

    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN" + "W.CWP")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN" + "W.CWP") ("$TMP_DIR\$global:FN" + "W.CWP")       
        }
    }
    
    WAVCOPY 
}

#WAVCOPY
function WAVCOPY {  
    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN?.WAV")) -eq $true) {
            $ARRAYS | foreach {
                if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN$_.WAV")) -eq $true) {
                    COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN$_.WAV") ("$TMP_DIR\$global:FN$_.WAV")
                }
            }
        }
    }   
    
    RBBCOPY
}

#:/ ----- Ver3.06(BB対応) 追加 ----------------------
#RBBCOPY
function RBBCOPY {    
   
    if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.RBB")) -eq $true) {
        COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.RBB") ("$TMP_DIR\$global:FN" + "W.RBB")             
    }

    RCCCOPY 
}

#:/ ----- Ver3.23(JS-W1対応) 追加 ----------------------
#RCCCOPY
function RCCCOPY {   
    if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.RCC")) -eq $true) {
        COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.RCC") ("$TMP_DIR\$global:FN" + "W.RCC")             
    }

    BMCCCOPY 
}

#BMCCCOPY
function BMCCCOPY {   
    if ((CheckExistPath("$global:KANRI_D\BMC\??$global:SLC*.BMC")) -eq $true) {
        COPY("$global:KANRI_D\BMC\??$global:SLC*.BMC") ("$TMP_DIR\$global:FN" + "W.BMC")             
    }

    DDCCOPY 
}

#DDCCOPY
function DDCCOPY { 

    if ((CheckExistPath("$global:KANRI_D\DDC\??$global:SLC*.DDC")) -eq $true) {
        COPY("$global:KANRI_D\DDC\??$global:SLC*.DDC") ("$TMP_DIR\$global:FN" + "W.DDC")             
    }

    AACCOPY 
}

#:/ ----- 以上、Ver3.23(JS-W1対応) 追加終了 ----------------------
#AACCOPY
function AACCOPY {   
    if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN" + "?.aac")) -eq $true) {
        $ARRAYS | foreach {
            if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN$_.aac")) -eq $true) {
                COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\$global:FN$_.aac") ("$TMP_DIR\$global:FN$_.aac")
            }
        }
    }
    
    GUIDECOPY 
}


#GUIDECOPY
function GUIDECOPY {  
    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {

        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_001.WAV")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_001.WAV") ("$TMP_DIR\$global:FN" + "_001.WAV")             
        }
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_002.WAV")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_002.WAV") ("$TMP_DIR\$global:FN" + "_002.WAV")             
        }
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_003.WAV")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_003.WAV") ("$TMP_DIR\$global:FN" + "_003.WAV")             
        }
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_004.WAV")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_004.WAV") ("$TMP_DIR\$global:FN" + "_004.WAV")             
        }
        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_099.WAV")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "_099.WAV") ("$TMP_DIR\$global:FN" + "_099.WAV")             
        }
    } 
    
    NAMAENSOU
}


#NAMAENSOU
function NAMAENSOU {
   
    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {

        if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "???.aac")) -eq $true) {
            COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "???.aac") ("$TMP_DIR")             
        }        
    }   
    
    REECOPY 
}


#REECOPY
function REECOPY {  

    if ((CheckExistPath("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.REE")) -eq $true) {
        COPY("$global:KANRI_D\MP2WAVE\data\$global:SLC\??$global:SLC" + "W.REE") ("$TMP_DIR\$global:FN" + "W.REE")             
    } 
    
    #:/／￣￣ PTNファイル準備 ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
    #:/＜処理説明＞
    #:/　PTNファイルを選択させ、自動でPTNファイルを作成する。
    #:/　PTNファイルが既に存在する場合は、処理しない。

    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.PTN")) -eq $false) {
        $pnt_selected_Index = ShowSelectionForm 

        #Check start Index = 1
        $pnt_selected_Index = $pnt_selected_Index + 1             

        if ($pnt_selected_Index -eq 0) {
            EXIT_PTNSET   
        }
        elseif ($pnt_selected_Index -eq 1) {
            COPY ("$EXE_DIR\RED-1.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET       
        }
        elseif ($pnt_selected_Index -eq 2) {
            COPY ("$EXE_DIR\RED-3.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET       
        }
        elseif ($pnt_selected_Index -eq 3) {
            COPY ("$EXE_DIR\GREEN-1.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET          
        }
        elseif ($pnt_selected_Index -eq 4) {
            COPY ("$EXE_DIR\GREEN-3.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET         
        }
        elseif ($pnt_selected_Index -eq 5) {
            COPY ("$EXE_DIR\BLUE-1.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET     
        }
        elseif ($pnt_selected_Index -eq 6) {
            COPY ("$EXE_DIR\BLUE-3.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET         
        }
        elseif ($pnt_selected_Index -eq 7) {
            COPY ("$EXE_DIR\POP-B.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET           
        }
        elseif ($pnt_selected_Index -eq 8) {
            COPY ("$EXE_DIR\POP-G.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET   
        }
        elseif ($pnt_selected_Index -eq 9) {
            COPY ("$EXE_DIR\POP-M.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET        
        }
        elseif ($pnt_selected_Index -eq 10) {
            COPY ("$EXE_DIR\POP-M2.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET      
        }
        elseif ($pnt_selected_Index -eq 11) {
            COPY ("$EXE_DIR\POP-P.PTN") ("$TMP_DIR\$global:FN" + "W.PTN")
            EXIT_PTNSET     
        }        
    }

    RCPREV
}

#EXIT_PTNSET
function EXIT_PTNSET {  
    RCPREV 
}

#RCPREV
function RCPREV { 
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        CPRCP_NEW
        return
    }

    $global:DRV = "$FD_DRV\88"
    $global:RN = "R36"

    #サブルーチンへ
    GET_RCPREV

    $global:R88_N = $global:RN
    $global:R88_REV = $global:REV

    if ((CheckExistPath("$FD_DRV\55\*.*")) -eq $true) {
        $global:DRV = "$FD_DRV\55"
        $global:RN = "R36"

        #サブルーチンへ
        GET_RCPREV

        $global:R55_N = $global:RN
        $global:R55_REV = $global:REV
    }

    #:/ ----- Ver3.05追加 ----------------------

    if ((CheckExistPath("$FD_DRV\AA\*.R36")) -eq $true) {
        $global:DRV = "$FD_DRV\AA"
        $global:RN = "R36"

        #サブルーチンへ
        GET_RCPREV

        $global:RAA_N = $global:RN
        $global:RAA_REV = $global:REV
    }

    #:/ ----- Ver3.06追加 ----------------------
    if ((CheckExistPath("$FD_DRV\BB\*.R36")) -eq $true) {
        $global:DRV = "$FD_DRV\BB"
        $global:RN = "R36"

        #サブルーチンへ
        GET_RCPREV

        $global:RBB_N = $global:RN
        $global:RBB_REV = $global:REV
    }

    #:/ ----- Ver3.23追加 ----------------------
    if ((CheckExistPath("$FD_DRV\CC\*.R36")) -eq $true) {
        $global:DRV = "$FD_DRV\CC"
        $global:RN = "R36"

        #サブルーチンへ
        GET_RCPREV

        $global:RCC_N = $global:RN
        $global:RCC_REV = $global:REV
    }

    CPRCP_NEW 
}

#:/／￣￣ ＦＤよりコピー ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
#:/ネットワーク環境、社内環境は、次の処理へ
#CPRCP_NEW
function CPRCP_NEW {  

    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        NET_RCP
        return
    }

    if ((CheckExistPath("$FD_DRV\55\*.R36")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.R55")) -eq $true) {
            DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "?.R55")
        }

        COPY ("$FD_DRV\55\??$global:SLC$global:R55_REV.R36") ("$TMP_DIR\$global:FN" + "W.R55")

        if ($? -eq $false) {
            $global:ERR_FNAME = "R55"
            CP_ERR
        }        
    }

    if ((CheckExistPath("$FD_DRV\88\??$global:SLC$global:R88_REV.R36")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.R88")) -eq $true) {       
            DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "?.R88")
        }
        COPY ("$FD_DRV\88\??$global:SLC$global:R88_REV.R36") ("$TMP_DIR\$global:FN" + "W.R88")

        if ($? -eq $false) {
            $global:ERR_FNAME = "R88"
            CP_ERR
        }        
    }

    #:/ ----- 以下、Ver3.05追加 ----------------------
    if ((CheckExistPath("$FD_DRV\AA\??$global:SLC$global:RAA_REV.R36")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.RAA")) -eq $true) {       
            DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "?.RAA")
        }

        COPY ("$FD_DRV\AA\??$global:SLC$global:RAA_REV.R36") ("$TMP_DIR\$global:FN" + "W.RAA")

        if ($? -eq $false) {
            $global:ERR_FNAME = "RAA"
            CP_ERR
        }       
    }

    #:/ ----- 以下、Ver3.06追加 ----------------------
    if ((CheckExistPath("$FD_DRV\BB\??$global:SLC$global:RBB_REV.R36")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.RBB")) -eq $true) {       
            DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "?.RBB")
        }

        COPY ("$FD_DRV\BB\??$global:SLC$global:RBB_REV.R36") ("$TMP_DIR\$global:FN" + "W.RBB")

        if ($? -eq $false) {
            $global:ERR_FNAME = "RBB"
            CP_ERR           
        }
    }

    #:/ ----- 以下、Ver3.23追加 ----------------------
    if ((CheckExistPath("$FD_DRV\CC\??$global:SLC$global:RCC_REV.R36")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.RCC")) -eq $true) {       
            DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "?.RCC")
        }

        COPY ("$FD_DRV\CC\??$global:SLC$global:RCC_REV.R36") ("$TMP_DIR\$global:FN" + "W.RCC")

        if ($? -eq $false) {
            $global:ERR_FNAME = "RCC"
            CP_ERR           
        }
    }
 
    NET_RCP
}

#/社外ネットワーク、社内環境
#NET_RCP
function NET_RCP {   

    ##2-4
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {       

        if ((CheckExistPath("$global:RCP_DIR\??$global:SLC" + "W.R55")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.R55")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "W.R55")               
            }
            COPY("$global:RCP_DIR\??$global:SLC" + "W.R55") ("$TMP_DIR\$global:FN" + "W.R55")
            if ($? -eq $false) {
                $global:ERR_FNAME = "R55"
                CP_ERR                
            }
        }

        if ((CheckExistPath("$global:RCP_DIR\??$global:SLC" + "W.R88")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.R88")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "W.R88")               
            }
            COPY("$global:RCP_DIR\??$global:SLC" + "W.R88") ("$TMP_DIR\$global:FN" + "W.R88")
            if ($? -eq $false) {
                $global:ERR_FNAME = "R88"
                CP_ERR                
            }
        }

        if ((CheckExistPath("$global:RCP_DIR\??$global:SLC" + "W.RAA")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.RAA")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "W.RAA")               
            }
            COPY("$global:RCP_DIR\??$global:SLC" + "W.RAA") ("$TMP_DIR\$global:FN" + "W.RAA")
            if ($? -eq $false) {
                $global:ERR_FNAME = "RAA"
                CP_ERR                
            }
        }

        if ((CheckExistPath("$global:RCP_DIR\??$global:SLC" + "W.RBB")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.RBB")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "W.RBB")               
            }
            COPY("$global:RCP_DIR\??$global:SLC" + "W.RBB") ("$TMP_DIR\$global:FN" + "W.RBB")
            if ($? -eq $false) {
                $global:ERR_FNAME = "RBB"
                CP_ERR                
            }
        }

        #Ver3.23追加
        if ((CheckExistPath("$global:RCP_DIR\??$global:SLC" + "W.RCC")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.RCC")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "W.RCC")               
            }
            COPY("$global:RCP_DIR\??$global:SLC" + "W.RCC") ("$TMP_DIR\$global:FN" + "W.RCC")
            if ($? -eq $false) {
                $global:ERR_FNAME = "RCC"
                CP_ERR                
            }
        }

        #Ver3.34追加
        if ((CheckExistPath("$global:RCP_DIR\??$global:SLC" + "W_R88.MID")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W_R88.MID")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "W_R88.MID")               
            }
            COPY("$global:RCP_DIR\??$global:SLC" + "W_R88.MID") ("$TMP_DIR\$global:FN" + "W_R88.MID")
            if ($? -eq $false) {
                $global:ERR_FNAME = "RCC"
                CP_ERR                
            }
        }

        #Ver3.40追加
        if ((CheckExistPath("$global:RCP_DIR\??$global:SLC" + "W.REE")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.REE")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\??$global:SLC" + "W.REE")               
            }
            COPY("$global:RCP_DIR\??$global:SLC" + "W.REE") ("$TMP_DIR\$global:FN" + "W.REE")
            if ($? -eq $false) {
                $global:ERR_FNAME = "REE"
                CP_ERR                
            }
        }

        #Ver4.03追加
        if ((CheckExistPath("$global:RCP_DIR\NJ0$global:SLC.MID")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\NJ0$global:SLC.MID")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\NJ0$global:SLC.MID")               
            }
            COPY("$global:RCP_DIR\NJ0$global:SLC.MID") ("$TMP_DIR\NJ0$global:SLC.MID")
            if ($? -eq $false) {
                $global:ERR_FNAME = "UGAMID"
                CP_ERR                
            }
        }

        #Ver4.03追加
        if ((CheckExistPath("$global:RCP_DIR\J0$global:SLC.MID")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\J0$global:SLC.MID")) -eq $true) {           
                DeleteFileOrFolder ("$TMP_DIR\J0$global:SLC.MID")               
            }
            COPY("$global:RCP_DIR\J0$global:SLC.MID") ("$TMP_DIR\J0$global:SLC.MID")
            if ($? -eq $false) {
                $global:ERR_FNAME = "UGAMID"
                CP_ERR                
            }
        }

        #####################START_INS_2_検査ソフトハウス#########################################
        if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 4) {
            #Ver4.66追加
            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_R02.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R02.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_R02.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_R02.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "R02"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_R02R88_EFXON_Track*.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R02R88_EFXON_Track*.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_R02R88_EFXON_Track*.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_R02R88_EFXON_Track*.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXMID"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_RB2R88_EFXON_Track*.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RB2R88_EFXON_Track*.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_RB2R88_EFXON_Track*.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_RB2R88_EFXON_Track*.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXMID"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_R02UGA_EFXON_Track*.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R02UGA_EFXON_Track*.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_R02UGA_EFXON_Track*.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_R02UGA_EFXON_Track*.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXMID"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_R88tmp_EFXOFF_Track*.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R88tmp_EFXOFF_Track*.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_R88tmp_EFXOFF_Track*.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_R88tmp_EFXOFF_Track*.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXMID"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_RAAtmp_EFXOFF_Track*.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RAAtmp_EFXOFF_Track*.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_RAAtmp_EFXOFF_Track*.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_RAAtmp_EFXOFF_Track*.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXMID"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_UGAtmp_EFXOFF_Track*.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_UGAtmp_EFXOFF_Track*.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_UGAtmp_EFXOFF_Track*.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_UGAtmp_EFXOFF_Track*.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXMID"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_R02R88_Track*.txt")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R02R88_Track*.txt")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_R02R88_Track*.txt")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_R02R88_Track*.txt") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXTXT"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_RB2R88_Track*.txt")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RB2R88_Track*.txt")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_RB2R88_Track*.txt")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_RB2R88_Track*.txt") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXTXT"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_R02UGA_Track*.txt")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R02UGA_Track*.txt")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_R02UGA_Track*.txt")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_R02UGA_Track*.txt") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "EFXTXT"
                    CP_ERR                
                }
            }

            if ((CheckExistPath("$global:RCP_DIR\$global:FN" + "W_RB2.mid")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RB2.mid")) -eq $true) {           
                    DeleteFileOrFolder ("$TMP_DIR\$global:FN" + "W_RB2.mid")               
                }

                COPY("$global:RCP_DIR\$global:FN" + "W_RB2.mid") ("$TMP_DIR")
                if ($? -eq $false) {
                    $global:ERR_FNAME = "RB2"
                    CP_ERR                
                }
            }
        }
        #####################END_INS_2_検査ソフトハウス#########################################
    }

    CPDB 
}

#:/Ver3.14 社内と社外ネットワークはFD_DRV→DB_DIRに変更
#:/社外FDはFDのまま
function CPDB {  
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) { 
        DB_COPY 
        return
    }

    if ((CheckExistPath("$FD_DRV\$global:FN" + "W.DBN")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DBN")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DBN")
        }
    
        COPY("$FD_DRV\$global:FN" + "W.DBN") ("$TMP_DIR\$global:FN" + "W.DBN")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DBN"
            CP_ERR                
        }
    }
    else {
        $global:ERR_FNAME = "DBN"
        NO_FILE_ERR   
    }
    
    if ((CheckExistPath("$FD_DRV\$global:FN" + "W.DBM")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DBM")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DBM")
        }
    
        COPY("$FD_DRV\$global:FN" + "W.DBM") ("$TMP_DIR\$global:FN" + "W.DBM")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DBM"
            CP_ERR                
        }
    }

    CPDDB 
}

#CPDDB
function CPDDB {          
    if ((CheckExistPath("$FD_DRV\$global:FN" + "W.DDB")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDB")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDB")
        }
    
        COPY("$FD_DRV\$global:FN" + "W.DDB") ("$TMP_DIR\$global:FN" + "W.DDB")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DDB"
            CP_ERR                
        }
    }

    CPDDA 
}

#:/ ----- Ver3.04(JS70対応) 追加 ----------------------
#CPDDA
function CPDDA {

    if ((CheckExistPath("$FD_DRV\$global:FN" + "W.DDA")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDA")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDA")
        }
    
        COPY("$FD_DRV\$global:FN" + "W.DDA") ("$TMP_DIR\$global:FN" + "W.DDA")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DDA"
            CP_ERR                
        }
    }

    CPDDC
}

#:/ ----- Ver3.23(JS-W1対応) 追加 ----------------------
#CPDDC
function CPDDC {
    if ((CheckExistPath("$FD_DRV\$global:FN" + "W.DDC")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDC")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDC")
        }
    
        COPY("$FD_DRV\$global:FN" + "W.DDC") ("$TMP_DIR\$global:FN" + "W.DDC")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DDC"
            CP_ERR                
        }
    }

    CPDDE
}

#DDE
function CPDDE {
    if ((CheckExistPath("$FD_DRV\$global:FN" + "W.DDE")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDE")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDE")
        }
    
        COPY("$FD_DRV\$global:FN" + "W.DDE") ("$TMP_DIR\$global:FN" + "W.DDE")
        if ($? -eq $false) {
            $global:ERR_FNAME = "DDE"
            CP_ERR                
        }
    }

    CPK3T
}

#CPK3T
function CPK3T {
    if ((CheckExistPath("$TMP_DIR\12$global:SLC.k3t")) -eq $true) {
        DeleteFileOrFolder("$TMP_DIR\12$global:SLC.k3t")
    }

    if ((CheckExistPath("$FD_DRV\12$global:SLC.k3t")) -eq $true) {
        COPY("$FD_DRV\12$global:SLC.K3t") ("$TMP_DIR")
        if ($? -eq $false) {
            $global:ERR_FNAME = "K3T"
            CP_ERR                
        }
    }
}

#DB_COPY
function DB_COPY {
  
    #:/ネットワーク環境、社内環境のみ
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) { 
        #Write-Host "DB_COPY path: $DB_DIR\$global:FN" + "W.DBN"
        #return 
        if ((CheckExistPath("$DB_DIR\$global:FN" + "W.DBN")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DBN")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DBN")
            }

            COPY("$DB_DIR\$global:FN" + "W.DBN") ("$TMP_DIR\$global:FN" + "W.DBN")
            if ($? -eq $false) {
                $global:ERR_FNAME = "DBN"
                CP_ERR                
            }
        }
        else {
            $global:ERR_FNAME = "DBN"
            NO_FILE_ERR   
        }

        #CPDDB
        if ((CheckExistPath("$DB_DIR\$global:FN" + "W.DDB")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDB")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDB")
            }
            COPY("$DB_DIR\$global:FN" + "W.DDB") ("$TMP_DIR\$global:FN" + "W.DDB")
            if ($? -eq $false) {
                $global:ERR_FNAME = "DDB"
                CP_ERR                
            }
        }

        #:/ ----- Ver3.04(JS70対応) 追加 ----------------------
        #CPDDA
        if ((CheckExistPath("$DB_DIR\$global:FN" + "W.DDA")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDA")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDA")
            }
            COPY("$DB_DIR\$global:FN" + "W.DDA") ("$TMP_DIR\$global:FN" + "W.DDA")
            if ($? -eq $false) {
                $global:ERR_FNAME = "DDA"
                CP_ERR                
            }
        }

        #:/ ----- Ver3.23(JS-W1対応) 追加 ----------------------
        #CPDDC
        if ((CheckExistPath("$DB_DIR\$global:FN" + "W.DDC")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDC")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDC")
            }
            COPY("$DB_DIR\$global:FN" + "W.DDC") ("$TMP_DIR\$global:FN" + "W.DDC")
            if ($? -eq $false) {
                $global:ERR_FNAME = "DDC"
                CP_ERR                
            }
        }

        #CPDDE
        if ((CheckExistPath("$DB_DIR\$global:FN" + "W.DDE")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DDE")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DDE")
            }
            COPY("$DB_DIR\$global:FN" + "W.DDE") ("$TMP_DIR\$global:FN" + "W.DDE")
            if ($? -eq $false) {
                $global:ERR_FNAME = "DDE"
                CP_ERR                
            }
        }

        #CPDBM
        if ((CheckExistPath("$DB_DIR\$global:FN" + "W.DBM")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "?.DBM")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\??$global:SLC" + "?.DBM")
            }
            COPY("$DB_DIR\$global:FN" + "W.DBM") ("$TMP_DIR\$global:FN" + "W.DBM")
            if ($? -eq $false) {
                $global:ERR_FNAME = "DBM"
                CP_ERR                
            }
        }

        #:/ ----- Ver4.30(XDB連携対応) 追加 ----------------------
        #CPSAD
       
        $bemax = 0
        $bemaxZero = 0
        $keta6bemax = 0

        if ((CheckExistPath("$DB_DIR\$global:SLC.BMN")) -eq $true) {

            $bmn = GetContentByIndex ("$DB_DIR\$global:SLC.BMN") 0

            #BeMAX番号をBMNから取得
            #Get substring by tab   
            $bemax = GetSubArrayInArrays $bmn "`t" 0 1
            #waitting
            $bemax = $bmn

            if ($bemax.Length -eq 1) {
                $bemaxZero = "0000$bemax"
            }
            elseif ($bemax.Length -eq 2) {
                $bemaxZero = "000$bemax"
            }
            elseif ($bemax.Length -eq 3) {
                $bemaxZero = "00$bemax"
            }
            elseif ($bemax.Length -eq 4) {
                $bemaxZero = "0$bemax"
            }
            elseif ($bemax.Length -eq 5) {
                $bemaxZero = "$bemax"
            }
            elseif ($bemax.Length -eq 6) {
                $bemaxZero = "$bemax"
                $keta6bemax = 1
            }
            else {
                SADCP_ERR
            }

            #SADをコピー
            ########################START_INS_1_XING社内#############################     
            if ($global:METHOD_TYPE -eq 4) {      
                if ($keta6bemax -eq 0) {
                    if ((CheckExistPath("$global:FLG_DIR\???$bemaxZero.SAD")) -eq $true) {
                        COPY ("$global:FLG_DIR\???$bemaxZero.SAD") ("$TMP_DIR")
                    }               
                }
                else {
                    if ((CheckExistPath("$global:FLG_DIR\??$bemaxZero.SAD")) -eq $true) {
                        COPY ("$global:FLG_DIR\??$bemaxZero.SAD") ("$TMP_DIR")
                    }            
                }
            }
            else {
                if ($keta6bemax -eq 0) {
                    if ((CheckExistPath("$DB_DIR\???$bemaxZero.SAD")) -eq $true) {
                        COPY ("$DB_DIR\???$bemaxZero.SAD") ("$TMP_DIR")
                    } 
                }
                else {
                    if ((CheckExistPath("$DB_DIR\??$bemaxZero.SAD")) -eq $true) {
                        COPY ("$DB_DIR\??$bemaxZero.SAD") ("$TMP_DIR")
                    }
                }             
            }
            ########################END_INS_2_検査ソフトハウス#############################
        }
        
    }  

    CSV_PROCESSING 
}

#CSV_PROCESSING
function CSV_PROCESSING {  

    #:/Network環境のみ実行
    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
        #:/ローカルにCSVファイルが2つあれば、合体させる
      
        if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.CSV")) -eq $true) {
            if ((CheckExistPath("$TMP_DIR\get.csv")) -eq $true) {
                #Write-Host "CSV_PROCESSING -test1"
                #発注区分取得
                SaveFileInfo ("$TMP_DIR\??$global:SLC" + "W.CSV") ("$TMP_DIR\a.txt")
                
                $sh_tmp = GetContentByIndex ("$TMP_DIR\a.txt") 0

                $global:sh = $sh_tmp.Substring(0, 2)
                DeleteFileOrFolder("$TMP_DIR\a.txt")

                if ($? -eq $false) {
                    $global:ERR_FNAME = "CSV"
                    CP_ERR
                }

                if ((CheckExistPath("$TMP_DIR\??$global:SLC" + "W.R88")) -eq $true) {
                    #/イントロタイムを取得                   
                    $global:INTRO = ""
                    $par1 = "$TMP_DIR\$global:SLC$global:SLC" + "W.R88"
                    $par2 = "$EXE_DIR\getintro.txt"

                    cmd.exe /c $GET_INTRO_TIME_FROM_RCP_EXE_PATH $par1> $par2 

                    #Write-Host "CSV_PROCESSING -test2"

                    if ($LASTEXITCODE -eq -1) {
                        $global:INTRO = ""
                    }
                    else {                        
                        $global:INTRO = $LASTEXITCODE
                    }
                }
                else {
                    $global:INTRO = ""
                }
                
                $args1 = "$TMP_DIR\get.csv"
                $args2 = "$TMP_DIR\$sh$global:SLC" + "W.CSV"
                $args3 = $global:INTRO
                
                cmd.exe /c $UNION_CSV_EXE_PATH $args1 $args2 $args3

                DeleteFileOrFolder("$TMP_DIR\get.csv")
                #Write-Host "CSV_PROCESSING -test3"
            }
        }
    }

    #:/／￣￣ 各種フォーマットチェック ￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣
    #:/＜処理説明＞
    #:/　PLY,RCPのフォーマットチェック、デュエット区分の取得を行う。
    #:/ ----- Ver3.02 追加----------------------
    #:/ PLY Format Check & Color Pallet変換
    if ((CheckExistPath("$TMP_DIR\rcpchks.err")) -eq $true) {
        DeleteFileOrFolder("$TMP_DIR\rcpchks.err")
    }
    
    #:/(Ver3.19変更)デュエット区分取得をVBSから行う 
    $d_type = ""

    #:/VBSで読み取れるよう、CSVにリネームする   

    COPY("$TMP_DIR\$global:FN" + "W.DBN") ("$TMP_DIR\getdbn.csv")

    if ((CheckExistPath("$GET_DBN2_EXE_PATH")) -eq $true) {
        #:/Ver3.32処理変更  
        cmd.exe /c $GET_DBN2_EXE_PATH "$TMP_DIR\getdbn.csv" "$TMP_DIR\getdbn.txt"
    }
    else {
        ShowMessageError "Error" "Can't found file $GET_DBN2_EXE_PATH to run!!"
    }

    $duet_tmp = [string] (GetContentByIndex ("$TMP_DIR\getdbn.txt") 0)
    
    #:/デュエット曲だった場合はRCPchksの引数に「duet」を加える
    if ($duet_tmp -eq "男男" -or $duet_tmp -eq "女女" -or $duet_tmp -eq "男女") {
        $d_type = $duet_tmp
    }

    DeleteFileOrFolder("$TMP_DIR\getdbn.csv")
   
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.R55")) -eq $true) {        
        $arg1 = "$global:FN" + "W"    
        
        cmd.exe /c $RCP_CHKS_EXE_PATH $arg1 R55 $d_type

        if ((CheckExistPath("$TMP_DIR\rcpchks.err")) -eq $true) {
            ShowMessageError "RCP Format or Compare Error!!!" "RCP Format or Compare Error!!!`n Please look [c:\\cms5\\data\\RCPChk.log] File."
        }
    }
    else {      
        $agr_dtype = "$global:FN" + "W"  
        cmd.exe /c  $RCP_CHKS_EXE_PATH $agr_dtype R88 $d_type

        if ((CheckExistPath("$TMP_DIR\rcpchks.err")) -eq $true) {
            ShowMessageError "RCP Format or Compare Error!!!" "RCP Format or Compare Error!!!`n Please look [c:\\cms5\\data\\RCPChk.log] File."
        }  
    }
   
    #:/Ver3.14社外ネットワーク環境のみ
    if ($global:METHOD_TYPE -eq 2 -or $global:METHOD_TYPE -eq 3 -or $global:METHOD_TYPE -eq 4 -or $global:METHOD_TYPE -eq 5 -or $global:METHOD_TYPE -eq 6) {
       
        #:/Ver3.17 Technical対応
        if ((CheckExistPath("$TMP_DIR\Technical.txt")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\Technical.txt")
        }
        else {
            #:/正常終了の場合のみFLGフォルダからFLG削除
            #:/Ver4.41 FrenchToast処理時はFLGの削除は不要
            if ($global:METHOD_TYPE -eq 6) {
            }
            else {
                DeleteFileOrFolder("$global:FLG_DIR\$global:FLG_NAME.FLG")
            }
        }
    }
    
    #:/Ver3.41追加　JD5のパラメータをJD6に自動修正(JD7があったら処理しない）
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD7")) -eq $true) {
    }
    elseif ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD5")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD6")) -eq $true) {

            if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD5_TMP")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.JD5_TMP")
            }            
            $par1 = "$TMP_DIR\$global:FN" + "W.JD5"
            $par2 = "$TMP_DIR\$global:FN" + "W.JD6"
            $par3 = "$TMP_DIR\$global:FN" + "W.JD5_TMP"
            $par4 = "$EXE_DIR\JD5.TXT"   
            if ((CheckExistPath("$JD5_PAR_SET_JD6_PATH")) -eq $true) {     
                cmd.exe /c $JD5_PAR_SET_JD6_PATH $global:SLC $par1 $par2 $par3>>$par4
            }
            else {
                ShowMessageError "Error" "Can't found file $JD5_PAR_SET_JD6_PATH to run!"
            }
                        
            #/書き換えが成功したら、書き換え後のJD5を真とする
            if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD5_TMP")) -eq $true) {
                DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.JD5")

                COPY("$TMP_DIR\$global:FN" + "W.JD5") ("$TMP_DIR\$global:FN" + "W.JD5")

                DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.JD5_TMP")
            }
            else {
                JD5CHG_ERR
            }
        }
    }

    RE_NAME_FILE
}

#VH区分の場合、VSからリネームして持ってくる
#RE_NAME_FILE
function RE_NAME_FILE { 
    if ($global:METHOD_TYPE -eq 0) {
        if ($global:FN_HEAD -eq "VH") {
            if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.DBM")) -eq $true) {
                $slc_tmp = ""

                $all_data = [string] (GetContentByIndex ("$TMP_DIR\$global:FN" + "W.DBM") 0)
               
                #選曲番号を5 or 6桁に変換
                $tmp = GetSubArrayInArrays $alldata "," 7 1

                #test
                ShowMessageError "ERR" "errr:$tmp"

                if ($tmp.Length -eq 1) {
                    $slc_tmp = "0000$tmp"
                }
                elseif ($tmp.Length -eq 2) {
                    $slc_tmp = "000$tmp"
                }
                elseif ($tmp.Length -eq 3) {
                    $slc_tmp = "00$tmp"
                }
                elseif ($tmp.Length -eq 4) {
                    $slc_tmp = "0$tmp"
                }
                elseif ($tmp.Length -eq 5) {
                    $slc_tmp = "$tmp"
                }
                elseif ($tmp.Length -eq 6) {
                    $slc_tmp = "$tmp"
                }
                else {
                    INPUT_ERROR
                }

                $BK_DIR_tmp = [String] (GetFolderNumber($slc_tmp))

                if ($BK_DIR_tmp.Length -eq 3) {
                    $BK_DIR_tmp = "00$BK_DIR_tmp"
                }
                elseif ($BK_DIR_tmp.Length -eq 4) {
                    $BK_DIR_tmp = "0$BK_DIR_tmp"
                }
                elseif ($BK_DIR_tmp.Length -eq 5) {                    
                }
                elseif ($BK_DIR_tmp.Length -eq 6) {     

                    $global:re = $BK_DIR_tmp.Substring(0, 1)
                    
                    switch ($global:re) {
                        1 {
                            $global:re = "_10\"
                            break
                        }
                        2 {
                            $global:re = "_20\"
                            break
                        }
                        3 {
                            $global:re = "_30\"
                            break
                        }
                        4 {
                            $global:re = "_40\"
                            break
                        }
                        5 {
                            $global:re = "_50\"
                            break
                        }
                        Default {}
                    }

                    $BK_DIR = "$global:re$BK_DIR_tmp"                  
                }
                else {
                    ShowMessageError "ERROR" "Back Up Directory Name Error!!"
                    $global:COMP_ERR = 1 
                    THE_END   
                }

                #元選曲番号のJD5をローカルに解凍
                if ((CheckExistPath("$BK_DRV\$BK_DIR\7WK$slc_tmp.LZH")) -eq $true) {

                    COPY ("$BK_DRV\$BK_DIR\7WK$slc_tmp.LZH") ("$TMP_DIR\$slc_tmp.LZH")
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "7WKLZH" 
                        CP_ERR 
                    }

                    $agr_jd7 = "$TMP_DIR\7WK$slc_tmp", "$TMP_DIR\ *.JD7"
                    RUN_EXE $WBLHA_EXE_PATH $agr_jd7

                    #発注区分を取得
                    SaveFileInfo ("$TMP_DIR\??$slc_tmp" + "W.JD7") ("$TMP_DIR\old_jd5.txt")

                    $JD5_NAME_tmp = (GetContentByIndex ("$TMP_DIR\old_jd5.txt") 0)

                    $sh_tmp = $JD5_NAME_tmp.Substring(0, 2)

                    #今あるVHのJD7を削除
                    DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.JD7")

                    #VSをVHにする
                    COPY ("$TMP_DIR\$sh_tmp\$slc_tmp" + "W.JD7") ("$TMP_DIR\$global:FN" + "W.JD7")
                    
                    $par1 = "$TMP_DIR\$global:FN" + "W.JD7"
                    $par2 = "$sh_tmp\$slc_tmp"
                    $par3 = $global:FN

                    if ((CheckExistPath("$REPLACE_CHAR_DATA_EXE_PATH")) -eq $true) {
                        cmd.exe /c $REPLACE_CHAR_DATA_EXE_PATH $par1 $par2 $par3 1
                    }
                    else {               
                        ShowMessageError "Error" "Can't found $REPLACE_CHAR_DATA_EXE_PATH to run!!!"
                    }

                    DeleteFileOrFolder("$TMP_DIR\7WK$slc_tmp.LZH")

                    DeleteFileOrFolder("$TMP_DIR\$sh_tmp$slc_tmp" + "W.JD7")                    
                }
                elseif ((CheckExistPath("$BK_DRV\$BK_DIR\5WK$slc_tmp.LZH")) -eq $true) {

                    COPY ("$BK_DRV\$BK_DIR\5WK$slc_tmp.LZH") ("$TMP_DIR\5WK$slc_tmp.LZH")
                    if ($? -eq $false) {
                        $global:ERR_FNAME = "5WKLZH" 
                        CP_ERR 
                    }

                    $agr_5WK = "E $TMP_DIR\5WK$slc_tmp $TMP_DIR\ *.JD5"

                    RUN_EXE $UNLHA_EXE_PATH $agr_5WK

                    #発注区分を取得

                    SaveFileInfo ("$TMP_DIR\??$slc_tmp" + "W.JD5") ("$TMP_DIR\old_jd5.txt")

                    $JD5_NAME = (GetContentByIndex ("$TMP_DIR\old_jd5.txt") 0)

                    $sh_tmp = $JD5_NAME.Substring(0, 2)

                    #今あるVHのJD5を削除
                    DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.JD5")

                    #VSをVHにする
                    COPY ("$TMP_DIR\$sh_tmp$slc_tmp" + "W.JD5") ("$TMP_DIR\5WK$global:FN" + "W.JD5")

                    $par1 = "$TMP_DIR\$global:FN" + "W.JD5"
                    $par2 = "$sh_tmp\$slc_tmp"
                    $par3 = $global:FN
                    
                    if ((CheckExistPath("$REPLACE_CHAR_DATA_EXE_PATH")) -eq $true) {
                        cmd.exe /c $REPLACE_CHAR_DATA_EXE_PATH $par1 $par2 $par3
                    }
                    else {               
                        ShowMessageError "Error" "Can't found $REPLACE_CHAR_DATA_EXE_PATH to run!!!"
                    }

                    DeleteFileOrFolder("$TMP_DIR\5WK$slc_tmp.LZH")
                    DeleteFileOrFolder("$TMP_DIR\$sh_tmp$slc_tmp" + "W.JD5")
                }
            }
        }
    }
    
    PLA_COL_PALLET 
}

#PLAColPalletChg
function PLA_COL_PALLET {     

    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD7")) -eq $true) {       
        $file_input = "$TMP_DIR\$global:FN" + "W.JD7"
        $file_output = "$TMP_DIR\a.JD7"

        if ((CheckExistPath("$PLA_COL_PALLET_CHG_EXE_PATH")) -eq $true) {  
            cmd.exe /c $PLA_COL_PALLET_CHG_EXE_PATH  $file_input > $file_output  
        }
        else {
            ShowMessageError "Error" "Can't found $PLA_COL_PALLET_CHG_EXE_PATH to run!!!"
        }

        DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.JD7")       
        COPY("$TMP_DIR\a.JD7") ("$TMP_DIR\$global:FN" + "W.JD7")

        DeleteFileOrFolder("$TMP_DIR\a.JD7")
    }

    OUTPUT_CHG 
}

#OUTPUT_CHG
function OUTPUT_CHG {   

    #################BEGIN_INS_2_検査ソフトハウス##################################
    if ($global:METHOD_TYPE -eq 4) {       
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD7")) -eq $true) {
            $agr_JD7 = "$TMP_DIR\$global:FN" + "W.JD7"  
            RUN_EXE $JD7_DVTON_EXE_PATH $agr_JD7
        } 
    }
    #####################END_INS_2_検査ソフトハウス##############################

    UGA_TITLE_WORDS_TO_BCHG 
}

#UGA_TITLE_WORDS_TO_BCHG
function UGA_TITLE_WORDS_TO_BCHG { 
   
    #################BEGIN_INS_1_XING社内##################################
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 4) {

        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD7")) -eq $true) {   
            if ((CheckExistPath("$UGA_TITLE_WORDS_TO_BCHG_EXE_PATH")) -eq $true) {  
                $file_input = "$TMP_DIR\$global:FN" + "W.JD7"
                $file_output = "$TMP_DIR\a.JD7"
                cmd.exe /c $UGA_TITLE_WORDS_TO_BCHG_EXE_PATH  $file_input > $file_output  

                if ($? -eq $false) {
                    $Error_Message = $Error[0]
                    ShowMessageError "Error - UGA_TITLE_WORDS_TO_BCHG" "Errors: $Error_Message"
                }
            }
            else {
                ShowMessageError "Error - UGA_TITLE_WORDS_TO_BCHG" "Can't found $UGA_TITLE_WORDS_TO_BCHG_EXE_PATH to run!!!"
            }

            DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.JD7")
            COPY("$TMP_DIR\a.JD7") ("$TMP_DIR\$global:FN" + "W.JD7")
            DeleteFileOrFolder("$TMP_DIR\a.JD7")           
        }
    }

    #################END_INS_1_XING社内##################################
    COMPARISON_JD7_AND_MID  
}

#COMPARISON_JD7_AND_MID 
function COMPARISON_JD7_AND_MID {
       
    DeleteFileOrFolder("$EXE_DIR\CompareJD7MID.TXT")
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD7")) -eq $true -and (CheckExistPath("$TMP_DIR\$global:FN" + "W_R88.mid")) -eq $true) {
        $par1 = "$TMP_DIR\$global:FN" + "W_R88.mid"
        $par2 = "$TMP_DIR\$global:FN" + "W.JD7"
        $par3 = "$TMP_DIR\$global:FN" + "W.DBN"
        $par4 = "$EXE_DIR\CompareJD7MID.TXT"

        if ((CheckExistPath("$COMPARE_JD7MID_EXE_PATH")) -eq $true) {
            cmd.exe /c $COMPARE_JD7MID_EXE_PATH $par1 $par2 $par3 > $par4
        }
        else {
            ShowMessageError "Error - COMPARISON_JD7_AND_MID" "Can't found $COMPARE_JD7MID_EXE_PATH to run!!!"
        }       
       
        if ($? -eq $false) {
            DeleteFileOrFolder("$EXE_DIR\CompareJD7MID.TXT")
        }
    }

    JD5_CHECK_ADDITION_DDB 
}

#Ver3.03 追加
#JD5 Check ※DDBの追加など
#JD5_CHECK_ADDITION_DDB
function JD5_CHECK_ADDITION_DDB {           

    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD5")) -eq $true) {
        #Ver3.40追加DDB挿入（以下3行）
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.DDB")) -eq $true) {
            if ((CheckExistPath("$INPUR_DDB_EXE_PATH")) -eq $true) {
                $par1 = "$TMP_DIR\$global:FN" + "W.JD5"
                $par2 = "$TMP_DIR\$global:FN" + "W.DDB"
                cmd.exe /c $INPUR_DDB_EXE_PATH $par1 $par2           
            }
            else {
                ShowMessageError "Error - JD5_CHECK_ADDITION_DDB" "Can't found $INPUR_DDB_EXE_PATH to run!!!"
            }     
        }

        #以上、Ver3.40追加       
        if ((CheckExistPath("$JD5_CHECK_EXE_PATH")) -eq $true) {
            $par1 = "$TMP_DIR\$global:FN" + "W.JD5"
            cmd.exe /c $JD5_CHECK_EXE_PATH $par1 
        }
        else {
            ShowMessageError "Error - JD5_CHECK_ADDITION_DDB" "Can't found $JD5_CHECK_EXE_PATH to run!!!"
        }  
    }

    if ((CheckExistPath("$TMP_DIR\$global:FN" + "D.JD5")) -eq $true) {

        if ((CheckExistPath("$JD5_CHECK_EXE_PATH")) -eq $true) {
            $par1 = "$TMP_DIR\$global:FN" + "D.JD5"
            cmd.exe /c $JD5_CHECK_EXE_PATH $par1 
        }
        else {
            ShowMessageError "Error - JD5_CHECK_ADDITION_DDB" "Can't found $JD5_CHECK_EXE_PATH to run!!!"
        }  
    }
    
    SMF_TO_RCP_RCP_TO_SMF 
}

#SMF_TO_RCP_RCP_TO_SMF 
function SMF_TO_RCP_RCP_TO_SMF {   

    #JD7が存在している場合のみ実行
    if ($global:METHOD_TYPE -eq 0 -or $global:METHOD_TYPE -eq 4 ) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD7")) -eq $true) {

            if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R88.MID")) -eq $true) {

                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.R88")) -eq $true) {
                    DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W.R88")
                }

                Set-Location $TMP_DIR
                
                if ((CheckExistPath("$SMF_TO_RCP_EXE")) -eq $true) {
                    #SMF to RCP 実行              
                    $par = "$TMP_DIR\$global:FN" + "W_R88.MID"
                    cmd.exe /c $SMF_TO_RCP_EXE  $par -A -L1                   
 
                    if ($? -eq $false) {
                        ShowMessageError "SMF to RCP Error!!!" "Don't convert SMF to RCP!!!`n Please look $SMF_TO_RCP_LOG. /E $SMF_TO_RCP_LOG"
                        THE_END
                    }
                    else {
                        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.r88")) -eq $true) {
                            $oldName = "$TMP_DIR\$global:FN" + "W.r88"
                            $newName = "$TMP_DIR\$global:FN" + "W.R88"                     
                            ReName $oldName $newName
                        }
                    }      
                }
                else {
                    ShowMessageError "Error - SMF_TO_RCP_RCP_TO_SMF" "Can't found $SMF_TO_RCP_EXE to run!!!"
                }          
            }

            if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.R88")) -eq $true) {
                if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_R88.mid")) -eq $true) {
                    DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W_R88.mid")
                }

                if ((CheckExistPath("$RCP_TO_SMF_EXE")) -eq $true) {
                    #RCP to SMF 実行              
                    $agr1 = "$TMP_DIR\$global:FN" + "W.R88"
                    $arg2 = "$TMP_DIR\$global:FN" + "W_R88.mid"                
                    cmd.exe /c $RCP_TO_SMF_EXE -A -L1 $agr1 -N:$arg2     

                    if ($? -eq $false) {
                        ShowMessageError "RCP to SMF Error!!!" "Don't convert RCP to SMF!!!`n Please look $RCP_TO_SMF_LOG. /E $RCP_TO_SMF_LOG"
                        THE_END
                    }
                }
                else {
                    ShowMessageError "Error - SMF_TO_RCP_RCP_TO_SMF" "Can't found $RCP_TO_SMF_EXE to run!!!"
                }   
            }            
        }
    }

    #:/Ver 4.10追加
    #:/RAA.mid,RBB.midの作成
    #:/ RAA
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.RAA")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RAA.mid")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W_RAA.mid")           
        }
       
        $par1 = "$TMP_DIR\$global:FN" + "W.RAA"
        $par2 = "$TMP_DIR\$global:FN" + "W_RAA.mid"
    
        cmd.exe /c $RCP_TO_SMF_EXE -A -L1 $par1 -N:$par2  

        if ($? -eq $false) {
            ShowMessageError "RCP to SMF Error!!!" "Don't convert RCP to SMF!!!`n Please look $RCP_TO_SMF_LOG. /E $RCP_TO_SMF_LOG"
            THE_END
        }          
    }

    #:/ RBB
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.RBB")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RBB.mid")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W_RBB.mid")           
        }

        $par1 = "$TMP_DIR\$global:FN" + "W.RBB"
        $par2 = "$TMP_DIR\$global:FN" + "W_RBB.mid"

        cmd.exe /c $RCP_TO_SMF_EXE -A -L1 $par1 -N:$par2

        if ($? -eq $false) {
            ShowMessageError "RCP to SMF Error!!!" "Don't convert RCP to SMF!!!`n Please look $RCP_TO_SMF_LOG. /E $RCP_TO_SMF_LOG"
            THE_END
        }
    }

    #:/ RCC
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.RCC")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RCC.mid")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W_RCC.mid")           
        }

        $par1 = "$TMP_DIR\$global:FN" + "W.RCC"
        $par2 = "$TMP_DIR\$global:FN" + "W_RCC.mid"
        cmd.exe /c $RCP_TO_SMF_EXE -A -L1 $par1 -N:$par2

        if ($? -eq $false) {
            ShowMessageError "RCP to SMF Error!!!" "Don't convert RCP to SMF!!!`n Please look $RCP_TO_SMF_LOG. /E $RCP_TO_SMF_LOG"
            THE_END
        }
    }

    #:/ RDD
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.RDD")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_RDD.mid")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W_RDD.mid")           
        }

        $par1 = "$TMP_DIR\$global:FN" + "W.RDD"
        $par2 = "$TMP_DIR\$global:FN" + "W_RDD.mid"
        cmd.exe /c $RCP_TO_SMF_EXE -A -L1 $par1 -N:$par2

        if ($? -eq $false) {
            ShowMessageError "RCP to SMF Error!!!" "Don't convert RCP to SMF!!!`n Please look $RCP_TO_SMF_LOG. /E $RCP_TO_SMF_LOG"
            THE_END
        }
    }

    #:/ REE
    if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.REE")) -eq $true) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W_REE.mid")) -eq $true) {
            DeleteFileOrFolder("$TMP_DIR\$global:FN" + "W_REE.mid")           
        }

        $par1 = "$TMP_DIR\$global:FN" + "W.REE"
        $par2 = "$TMP_DIR\$global:FN" + "W_REE.mid"
        cmd.exe /c $RCP_TO_SMF_EXE -A -L1 $par1 -N:$par2

        if ($? -eq $false) {
            ShowMessageError "RCP to SMF Error!!!" "Don't convert RCP to SMF!!!`n Please look $RCP_TO_SMF_LOG. /E $RCP_TO_SMF_LOG"
            THE_END
        }
    }

    MOVE_BMN 
}

#MOVE_BMN
function MOVE_BMN { 

    ########################BEGIN_INS_2_検査ソフトハウス#################################
    if ($global:METHOD_TYPE -eq 4) {
        if ((CheckExistPath("$TMP_DIR\$global:SLC.BMN")) -eq $true) {
            COPY("$TMP_DIR\$global:SLC.BMN") ("$DB_DIR\$global:SLC.BMN")
        }
    }
    ########################END_INS_2_検査ソフトハウス#################################
    INS_END
}

#:/Ver3.27移動
#INS_END
function INS_END {
    if ($ACOMP_SWITCH -ne 1) {
        if ((CheckExistPath("$TMP_DIR\$global:FN" + "W.JD7")) -eq $true) {
            ShowMessageInfo "INS" "Install is Finish!!"
        }
        else {
            ShowMessageInfo "INS" "Install is Finish OLD Data!!"
        }

        if ((CheckExistPath("$EXE_DIR\CompareJD7MID.TXT")) -eq $true) {
            StartProcessWait "notepad.exe" "$EXE_DIR\CompareJD7MID.TXT"            
        }
    }

    THE_END
}

#End thread
function THE_END {
    #:/フロッピードライブ名
    $FD_DRV = ""
    #:/ＴＭＰファイルのディレクトリ名
    $TMP_DIR = ""
    $EXE_DIR = ""
    #:/ＴＭＰファイル名
    $TMP_NAME = ""

    #選曲番号
    
    $global:re = ""
    $global:KANRI_D = ""
    $global:BK_DRV = ""
    $global:RCP_DIR = ""
    $global:AAC_DIR = ""
    $global:KEN_DIR = ""
    $global:ZEUS_DIR = ""
    $global:FLG_DIR = ""
    $global:FLG_DIR = ""
    $global:BK_DIR = ""
    $global:SLC = ""
    $global:file_name = ""
    $global:PSLC = ""
    $global:ERR_FNAME = ""
    $global:FN = ""
    $global:FN_HEAD = ""
    $global:RN = ""
    $global:EXTENSION = ""
    $global:DRV = ""
    $global:R88_N = ""
    $global:R88_REV = ""
    $global:REV = ""
    $global:R55_N = ""
    $global:R55_REV = ""
    $global:RAA_N = ""
    $global:RAA_REV = ""
    $global:RBB_N = ""
    $global:RBB_REV = ""
    $global:RCC_N = ""
    $global:RCC_REV = ""
    $global:INTRO = ""
    $global:FLG_NAME = ""

    if ($ACOMP_SWITCH -ne 1) {
        $global:COMP_ERR = ""
        #:/Ver3.14追加
        $global:file_name = ""
        $DB_DIR = ""
    }
    else {
        if ($global:COMP_ERR -eq 1) {
            WriteToFile ("NG") ("C:\cms5\ng.tmp")
        }
    }

    if ($formWaiting.Visible -eq $true) {
        $formWaiting.Close()         
    } 

    exit

    $global:IsStopThread = $true
}

#RCP_LOOP_END
function RCP_LOOP_END {
    #:/＊RCP OLD-NAME
    if ((CheckExistPath("$global:DRV\??$global:SLC" + "W.$global:RN")) -eq $true) {
        if ((CheckExistPath("$global:DRV\??$global:SLC" + "F.$global:RN")) -eq $true) {
            $global:REV = "F"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "E.$global:RN")) -eq $true) {
            $global:REV = "E"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "D.$global:RN")) -eq $true) {
            $global:REV = "D"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "C.$global:RN")) -eq $true) {
            $global:REV = "C"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "B.$global:RN")) -eq $true) {
            $global:REV = "B"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "A.$global:RN")) -eq $true) {
            $global:REV = "A"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "Z.$global:RN")) -eq $true) {
            $global:REV = "Z"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "Y.$global:RN")) -eq $true) {
            $global:REV = "Y"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "X.$global:RN")) -eq $true) {
            $global:REV = "X"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "W.$global:RN")) -eq $true) {
            $global:REV = "W"
        }
    }
    #:/＊RCP NEW-NAME
    elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "G.$global:RN")) -eq $true) {
        
        if ((CheckExistPath("$global:DRV\??$global:SLC" + "S.$global:RN")) -eq $true) {
            $global:REV = "S"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "Q.$global:RN")) -eq $true) {
            $global:REV = "Q"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "P.$global:RN")) -eq $true) {
            $global:REV = "P"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "O.$global:RN")) -eq $true) {
            $global:REV = "O"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "N.$global:RN")) -eq $true) {
            $global:REV = "N"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "M.$global:RN")) -eq $true) {
            $global:REV = "M"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "L.$global:RN")) -eq $true) {
            $global:REV = "L"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "K.$global:RN")) -eq $true) {
            $global:REV = "K"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "J.$global:RN")) -eq $true) {
            $global:REV = "J"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "H.$global:RN")) -eq $true) {
            $global:REV = "H"
        }
        elseif ((CheckExistPath("$global:DRV\??$global:SLC" + "G.$global:RN")) -eq $true) {
            $global:REV = "G"
        }
    }
    else {
        NO_BASEFILE
    }    
}

#RCPリビジョン判別（サブルーチン）
#GET_RCPREV
function GET_RCPREV {    
    $global:IsStopThread = $false
    
    #$global:DRV ="C:\collection"
    #$global:RN="log"
    #:/＊RCP OLD-NAME
    #:/Ver3.12 追加------------------------------------------------------
    #:/フロッピー内のR??ファイル名を書き出す
   
    SaveFileInfo ("$global:DRV\*.$global:RN") ("$TMP_DIR\$TMP_NAME")

    $filepath = "$TMP_DIR\$TMP_NAME"
    $reader = New-Object IO.StreamReader $filepath
  
    while ($null -ne ($RCP_NAME = $reader.ReadLine()) -and $global:IsStopThread -eq $false) { 
        #ファイルの最後まで検索したら終了
        if ($RCP_NAME.Length -eq 2) {
            RCP_LOOP_END            
        }
        #ファイルが読み込めない場合はエラー
        elseif ($RCP_NAME.Length -eq 1) {
            NO_TMP_FILE_ERR
        }
        #Exit thread
        if ($global:IsStopThread -eq $true) {
            break
        }

        $rr_max = $RCP_NAME.Length;     
        $rr = 1

        #チルダをチェックする為のループ
        while ($global:IsStopThread -eq $false) {  

            #1文字ずつ取り出す

            $chiruda = $RCP_NAME.Substring($rr, $rr + 1)

            #チルダチェック
            #Compare strings at First position
            #STR NCMP "string1" "string2" position => return:0 => same 1=>  not same"            
            if ($chiruda.IndexOf("~") -eq 0) {
                #Ver3.10 以下３行追加
                CHIRUDA_ERROR
            }
            else {
                $rr++;                        
            }   
        }        
    }
}
#########################################END_FUNCTION########################################################

#########################################START_THREAD_MAIN######################################

#Main thread
function THREAD_MAIN {
    
    Clear-Host

    ValidInput

    CheckDefaultFolder
       
    StartLocation    
    
    CHKDRV    
}

#Show form waiting
function FormWaiting {
    
    $rs = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.Open()

    $contents = "INS[$VER]for Windows2000`nAuthoring Data Installation Batch`n(A: -> $TMP_DIR)`n`nPlease Wait..."	
  
    $Label = New-Object System.Windows.Forms.Label
    $formWaiting.Text = "WAITTING";
    $formWaiting.MaximizeBox = $False
    $formWaiting.FormBorderStyle = 'FixedDialog'
    $formWaiting.Size = New-Object System.Drawing.Size(440, 240)
    $formWaiting.Controls.Add($Label)
    $formWaiting.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Label.Text = $contents
    $label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)   
    $label.Location = New-Object System.Drawing.Point(12, 35)
    $label.Size = New-Object System.Drawing.Size(440, 80)     

    #stop button
    $stopButton = New-Object System.Windows.Forms.Button
    $stopButton.Location = New-Object System.Drawing.Point(324, 157)
    $stopButton.Size = New-Object System.Drawing.Size(88, 32)
    $stopButton.Text = 'STOP'   
    $stopButton.Add_Click( {
            if ($formInputNumber.Visible -eq $true) {
                $formInputNumber.Close()
            }   
            if ($formSelection.Visible -eq $true) {
                $formSelection.Close()
            }         

            $global:IsStopThread = $true 
        })  
   
    $formWaiting.Controls.Add($stopButton)

    $formWaiting.Add_Shown( { $formWaiting.Activate() })

    $rs.SessionStateProxy.SetVariable("formWaiting", $formWaiting)

    $p = $rs.CreatePipeline( { [void] $formWaiting.ShowDialog() })
   
    $p.Input.Close()
    $p.InvokeAsync()

    Start-sleep 1

    THREAD_MAIN

    if ($formWaiting.Visible -eq $true) {
        $formWaiting.Close()   
        $rs.Close()  
    }  
}


####RUN THREAD####.

FormWaiting

#########################################END_THREAD_MAIN########################################
