<#
.SYNOPSIS
    Mock "PROKAS" pharmacy cash register window (WinForms).
    Loads products from products_list.txt, lets the user add quantities or
    register new PZNs via a dialog, and signals the OCR reader on changes.
#>

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -Namespace ProkasMock -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetProcessDPIAware();
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern System.IntPtr GetConsoleWindow();
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@

[void][ProkasMock.Native]::SetProcessDPIAware()
[void][ProkasMock.Native]::ShowWindow([ProkasMock.Native]::GetConsoleWindow(), 0)
[System.Windows.Forms.Application]::EnableVisualStyles()

# ===========================================================================
# 1. Configuration & State
# ===========================================================================
$script:WindowTitle  = 'PROKAS Kasse (Mock) - Filiale Entlebuch'
$script:ProductsFile = Join-Path $PSScriptRoot 'products_list.txt'
$script:Catalog      = [ordered]@{}
$script:Total        = 0.0

$script:RescanEvent       = New-Object System.Threading.EventWaitHandle($false, 'AutoReset', 'Local\ProkasOcrRescan')
$script:ManualRescanEvent = New-Object System.Threading.EventWaitHandle($false, 'AutoReset', 'Local\ProkasOcrManualRescan')

$script:FontRegular = New-Object System.Drawing.Font('Segoe UI', 13)
$script:FontBold    = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$script:FontMono    = New-Object System.Drawing.Font('Consolas', 12)

# ===========================================================================
# 2. Data Persistence (products_list.txt <-> $script:Catalog)
# ===========================================================================
function Import-ProductCatalog {
    if (-not (Test-Path $script:ProductsFile)) {
        $defaults = @(
            '01126111|Paracetamol 500mg 16 Tablette|4.90|1',
            '04345084|Ibuprofen 400mg 20 Tablette|6.90|1',
            '00999848|Wundsalbe Dexpanthenol 20g|8.50|1'
        )
        [System.IO.File]::WriteAllLines($script:ProductsFile, $defaults, [System.Text.Encoding]::UTF8)
    }

    foreach ($line in [System.IO.File]::ReadAllLines($script:ProductsFile, [System.Text.Encoding]::UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split '\|', 4
        if ($parts.Count -lt 4) { continue }

        $pzn = $parts[0].Trim()
        $script:Catalog[$pzn] = @{
            Name  = $parts[1].Trim()
            Price = [double]::Parse($parts[2].Trim(), [System.Globalization.CultureInfo]::InvariantCulture)
            Qty   = [int]::Parse($parts[3].Trim(), [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
}

function Add-ProductToFile([string]$Pzn, [hashtable]$Product) {
    $priceText = $Product.Price.ToString('F2', [System.Globalization.CultureInfo]::InvariantCulture)
    $line      = "{0}|{1}|{2}|{3}`r`n" -f $Pzn, $Product.Name, $priceText, $Product.Qty
    [System.IO.File]::AppendAllText($script:ProductsFile, $line, [System.Text.Encoding]::UTF8)
}

# ===========================================================================
# 3. UI Control Helpers
# ===========================================================================
function New-Label([string]$Text, [int]$X, [int]$Y, [int]$Width = 300, $Font = $script:FontRegular) {
    New-Object System.Windows.Forms.Label -Property @{
        Text   = $Text
        Left   = $X
        Top    = $Y
        Width  = $Width
        Height = 30
        Font   = $Font
    }
}

function New-TextBox([int]$X, [int]$Y, [int]$Width, [string]$DefaultText = '') {
    New-Object System.Windows.Forms.TextBox -Property @{
        Left  = $X
        Top   = $Y
        Width = $Width
        Text  = $DefaultText
        Font  = $script:FontRegular
    }
}

function New-Button([string]$Text, [int]$X, [int]$Y, [int]$Width = 125) {
    New-Object System.Windows.Forms.Button -Property @{
        Text   = $Text
        Left   = $X
        Top    = $Y
        Width  = $Width
        Height = 36
        Font   = $script:FontRegular
    }
}

# ===========================================================================
# 4. Main Window Layout
# ===========================================================================
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$initW  = [Math]::Max(760, [int]($screen.Width  * 0.45))
$initH  = [Math]::Max(500, [int]($screen.Height * 0.50))

$form = New-Object System.Windows.Forms.Form -Property @{
    Text          = $script:WindowTitle
    ClientSize    = New-Object System.Drawing.Size($initW, $initH)
    MinimumSize   = New-Object System.Drawing.Size(720, 420)
    StartPosition = 'CenterScreen'
    BackColor     = [System.Drawing.Color]::White
}

$lblHeader   = New-Label 'PROKAS Apotheke Kasse' 20 15 400 $script:FontBold
$lblCustomer = New-Label 'Kunde: Laufkunde'      20 55 400
$lblStatus   = New-Label 'Status: BEREIT' ($initW - 310) 15 290 $script:FontBold
$lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
$lblStatus.Anchor    = 'Top, Right'

$lblPzn    = New-Label   'PZN:'       20 100 60
$txtPzn    = New-TextBox              85  97 160
$lblQty    = New-Label   'Menge:'    265 100 75
$txtQty    = New-TextBox             345  97  60 '1'
$btnAdd    = New-Button  'Erfassen'  425  95 120
$btnRescan = New-Button  'OCR Rescan' 560 95 140

$lstItems = New-Object System.Windows.Forms.ListBox -Property @{
    Left   = 20
    Top    = 150
    Width  = ($initW - 40)
    Height = ($initH - 240)
    Font   = $script:FontMono
    Anchor = 'Top, Bottom, Left, Right'
}

$lblTotal = New-Label 'Total: CHF 0.00' ($initW - 310) ($initH - 75) 290 $script:FontBold
$lblTotal.Anchor = 'Bottom, Right'

$form.AcceptButton = $btnAdd
$form.Controls.AddRange(@(
    $lblHeader, $lblStatus, $lblCustomer,
    $lblPzn, $txtPzn, $lblQty, $txtQty, $btnAdd, $btnRescan,
    $lstItems, $lblTotal
))


function Set-Status([string]$Message, [System.Drawing.Color]$Color) {
    $lblStatus.Text      = "Status: $Message"
    $lblStatus.ForeColor = $Color
}

function Invoke-OcrRescan([switch]$Manual) {
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    if ($Manual) {
        [void]$script:ManualRescanEvent.Set()
    }
    [void]$script:RescanEvent.Set()
}

function Update-ReceiptList {
    $lstItems.BeginUpdate()
    $lstItems.Items.Clear()
    $script:Total = 0.0

    foreach ($entry in $script:Catalog.GetEnumerator()) {
        $pzn      = $entry.Key
        $product  = $entry.Value
        $lineCost = $product.Price * $product.Qty
        $script:Total += $lineCost

        $rowText = '{0}  {1,-30} {2,2}x  CHF {3,6:N2}' -f $pzn, $product.Name, $product.Qty, $lineCost
        [void]$lstItems.Items.Add($rowText)
    }

    $lstItems.EndUpdate()
    $lblTotal.Text = 'Total: CHF {0:N2}' -f $script:Total
}

function Show-NewArticleDialog([string]$Pzn, [string]$ErrorMessage = '', [string]$DraftName = '', [string]$DraftPrice = '0.00') {
    $dlg = New-Object System.Windows.Forms.Form -Property @{
        Text            = "Neuer Artikel anlegen (PZN: $Pzn)"
        ClientSize      = New-Object System.Drawing.Size(440, 235)
        StartPosition   = 'CenterParent'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        BackColor       = [System.Drawing.Color]::White
    }

    $lName  = New-Label   'Artikelname:' 20 25 130
    $tName  = New-TextBox               155 22 260 $DraftName
    $lPrice = New-Label   'Preis (CHF):' 20 75 130
    $tPrice = New-TextBox               155 72 120 $DraftPrice

    $lError = New-Label $ErrorMessage 20 115 400
    $lError.ForeColor = [System.Drawing.Color]::DarkRed

    $btnOk     = New-Button 'Speichern' 155 160 125
    $btnCancel = New-Button 'Abbrechen' 290 160 125
    $btnOk.DialogResult     = [System.Windows.Forms.DialogResult]::OK
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.AddRange(@($lName, $tName, $lPrice, $tPrice, $lError, $btnOk, $btnCancel))
    [void]$dlg.Add_Shown({ $tName.Focus() })

    $accepted = ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK)
    $rawInput = if ($accepted) { @{ Name = $tName.Text; Price = $tPrice.Text } } else { $null }
    $dlg.Dispose()
    return $rawInput
}

function Read-ValidatedNewProduct([string]$Pzn, [int]$Qty) {
    $errorMsg   = ''
    $draftName  = ''
    $draftPrice = '0.00'

    while ($true) {
        $inputData = Show-NewArticleDialog $Pzn $errorMsg $draftName $draftPrice
        if ($null -eq $inputData) { return $null }

        $draftName  = $inputData.Name
        $draftPrice = $inputData.Price

        try {
            $cleanName  = $inputData.Name.Trim()
            $cleanPrice = [double]::Parse(($inputData.Price.Trim() -replace ',', '.'), [System.Globalization.CultureInfo]::InvariantCulture)

            if ([string]::IsNullOrWhiteSpace($cleanName)) {
                throw 'Bitte einen Artikelnamen eingeben.'
            }
            if ($cleanPrice -le 0.00) {
                throw "Preis muss gr$([char]0x00F6)sser als 0.00 CHF sein."
            }

            return @{ Name = $cleanName; Price = $cleanPrice; Qty = $Qty }
        } catch {
            $errorMsg = "Fehler: $($_.Exception.Message)"
        }
    }
}

function Add-Item([string]$Pzn, [int]$Qty) {
    if ([string]::IsNullOrWhiteSpace($Pzn)) {
        Set-Status 'FEHLER PZN' ([System.Drawing.Color]::DarkRed)
        return
    }

    if ($script:Catalog.Contains($Pzn)) {
        $script:Catalog[$Pzn].Qty += $Qty
        Update-ReceiptList
        Set-Status 'ERFASST' ([System.Drawing.Color]::DarkGreen)
        Invoke-OcrRescan
        return
    }

    $newProduct = Read-ValidatedNewProduct $Pzn $Qty
    if ($null -eq $newProduct) {
        Set-Status 'ABGEBROCHEN' ([System.Drawing.Color]::DarkOrange)
        return
    }

    $script:Catalog.Add($Pzn, $newProduct)
    Add-ProductToFile $Pzn $newProduct
    Update-ReceiptList
    Set-Status "PZN hinzugef$([char]0x00FC)gt" ([System.Drawing.Color]::DarkGreen)
    Invoke-OcrRescan
}


$btnAdd.Add_Click({
    $qty = 0
    if (-not [int]::TryParse($txtQty.Text.Trim(), [ref]$qty) -or $qty -lt 1 -or $qty -gt 99) {
        Set-Status 'FEHLER MENGE' ([System.Drawing.Color]::DarkRed)
        return
    }
    Add-Item $txtPzn.Text.Trim() $qty
    $txtPzn.Clear()
    $txtQty.Text = '1'
    $txtPzn.Focus()
})

$btnRescan.Add_Click({
    Set-Status 'MANUELLER RESCAN' ([System.Drawing.Color]::DarkBlue)
    Invoke-OcrRescan -Manual
})

Import-ProductCatalog
Update-ReceiptList
Set-Status 'BEREIT' ([System.Drawing.Color]::DarkGreen)
[void]$form.Add_Shown({ $txtPzn.Focus() })
[System.Windows.Forms.Application]::Run($form)
