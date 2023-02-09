param(
[Parameter(Mandatory=$false)][string]$Path,
[Parameter(Mandatory=$false)][bool]$RootFiles=$false,
[Parameter(Mandatory=$false)][bool]$Recurse=$false
)

$files = $null

$db = cat "$PSScriptRoot\config.json" | ConvertFrom-Json

"Creating Worker: WORD"
$wordCom = New-Object -ComObject Word.Application
"Creating Worker: EXCEL"
$excelCom = New-Object -ComObject Excel.Application
"Creating Worker: OUTLOOK"
$outlook = New-Object -comobject outlook.application


function testString {
param(
    [Parameter(Mandatory=$true)]$str    
    )
    $result = New-Object System.Collections.ArrayList
    foreach($r in $db){
        Select-String -InputObject $str -Pattern $r.regex -AllMatches | % {
            $o = @{
                "type" = $r.type
                "matches" = $_.Matches.Value
            }
            if($o.matches -ne $null){
                $result.Add($o) | Out-Null
            }
            #$r.type + " Found."
        }
    }
    if($result.Count -gt 0){
        Write-Output $result
    }
}

function convert-PDFtoText {
	param(
		[Parameter(Mandatory=$true)]$file,
        [Parameter(Mandatory=$true)][string]$itext
	)	
	Add-Type -Path "$itext"
    $pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "$file"
	for ($page = 1; $page -le $pdf.NumberOfPages; $page++){
		$text=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
		Write-Output $text
	}	
	$pdf.Close()
}

function convert-WORDtoText {
	param(
		[Parameter(Mandatory=$true)]$file
	)

	$word = $wordCom.Documents.Open($file, $false, $true)
    Write-Output $word.content.Text
	$word.close()
}

function convert-ExceltoText {
	param(
		[Parameter(Mandatory=$true)]$file,
        [Parameter(Mandatory=$true)][int]$width,
        [Parameter(Mandatory=$true)][int]$height
	)	
	$excel = $excelCom.Workbooks.Open($file)
    $result = ''
    foreach($sheet in $document.Sheets){
        for($h = 1;$h -le $height;$h++){
            for($w = 1;$w -le $width;$w++){
                if($sheet.Rows($h).Cells($w).Text -ne ''){
                    $result += ($sheet.Rows($h).Cells($w).Text + ' ')
                }
            }
        }
    }

    Write-Output $result
	$excel.close()
}

function convert-MSGtoText {
	param(
		[Parameter(Mandatory=$true)]$file
	)

	$msg = $outlook.Session.OpenSharedItem($file)
    Write-Output $msg.body 
    $msg.Close(1)
}

$foundList = New-Object System.Collections.Generic.List[Object]

$root = ls -Path $Path -Directory

$root | % {
    "Searching Directory: $_"
    $dir = ls $_.FullName -Recurse

    foreach($item in $dir){
        "Reading File: " + $item.FullName        
        $str = $null

        # We don't want to try to open a directory.
        if($item.Attributes -ne 'Directory'){   
        
            # Switch for known filetype handlers    
            Switch -Wildcard ($item.FullName){
                "*.pdf" { $str = (convert-PDFtoText -file $_ -itext "$PSScriptRoot\itextsharp.dll") }
                "*.doc"{ $str = (convert-WORDtoText -file $item.FullName) }
                "*.docx"{ $str = (convert-WORDtoText -file $item.FullName) }
                "*.xlsx"{ $str = (convert-ExceltoText -file $item.FullName -height $hi -width $wi) }
                "*.xls"{ $str = (convert-ExceltoText -file $item.FullName -height $hi -width $wi) }
                "*.msg"{ $str = (convert-WORDtoText -file $item.FullName) }
                Default { $str = cat $item.FullName }
            }

            $search = testString -str $str
            if($search.Count -gt 0){
                # If the tester found matches, add them to the found list.
                $foundList.Add(@{file=$item.Name;path=$item.FullName;results=(testString -str $str)})
                $search | %{ "Found " + $_.type }
            }
        }
        
    }
}

$fileDate = (Date).ToString("yyyyMMDD")

if($foundList.Count -gt 0)
{
    "RUH-RO RAGGY!"
    "Hold on to your butts..."
    ConvertTo-Json $foundList -Depth 6 | Out-File ($PSScriptRoot + '\Results\' + $fileDate + '_Results.json')   
} else {
    "#!#!#!#!#!#! ALL CLEAR !#!#!#!#!#!#"
}

start $PSScriptRoot



$excelCom.Quit()
$outlook.Quit()
$wordCom.Quit()

