Add-Type -AssemblyName System.Windows.Forms

$SCRIPT:CONTROLS = @()

$SCRIPT:DEFAULT_FONT                 = [System.Drawing.Font]::New("Arial", 20, [System.Drawing.FontStyle]::Bold)
$SCRIPT:DEFAULT_COLOR_FOREGROUND     = [System.Drawing.Color]::FromArgb(255, 232, 230, 227)
$SCRIPT:DEFAULT_COLOR_BACKGROUND     = [System.Drawing.Color]::FromArgb(255, 14, 15, 16)
$SCRIPT:DEFAULT_COLOR_WRONG_SPOT     = [System.Drawing.Color]::FromArgb(255, 145, 127, 47)
$SCRIPT:DEFAULT_COLOR_NOT_IN_WORD    = [System.Drawing.Color]::FromArgb(255, 44, 48, 50)
$SCRIPT:DEFAULT_COLOR_CORRECT_SPOT   = [System.Drawing.Color]::FromArgb(255, 66, 113, 62)

$SCRIPT:WebRequest                   = Invoke-WebRequest -Uri "https://www.nytimes.com/games/wordle/index.html"
$SCRIPT:JS                           = Invoke-WebRequest -Uri ($SCRIPT:WebRequest.ParsedHtml.getElementsByTagName('script') | ? {$_.src -match "/games-assets/v2/wordle"}).src
$SCRIPT:WORDLIST                     = ($SCRIPT:JS.RawContent | sls "(?<=ia=\[)(.*?)(?=\])").Matches.Value -split ',' -replace '"'

Function Enumerate-Controls($Object) {
    foreach ($Control in $Object.Controls) {
        $SCRIPT:CONTROLS += $Control
        Enumerate-Controls $Control
    }
}

Function Create-RichTextBox {
    param(
        $Name        = "",
        $PosX        = 0,
        $PosY        = 0,
        $Width       = 62,
        $Height      = 62,
        $Enabled     = $true,
        $ForeColor   = $SCRIPT:DEFAULT_COLOR_FOREGROUND,
        $BackColor   = $SCRIPT:DEFAULT_COLOR_BACKGROUND,
        $TabIndex
    )

    # ===========================================================================================
    $RichTextBox                      = New-Object System.Windows.Forms.Label
    #============================================================================================

    $RichTextBox.Name            = $Name
    $RichTextBox.Font            = [System.Drawing.Font]::New("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $RichTextBox.Location        = New-Object System.Drawing.Point($PosX, $PosY)
    $RichTextBox.Width           = $Width
    $RichTextBox.Height          = $Height
    $RichTextBox.Enabled         = $Enabled
    $RichTextBox.ForeColor       = $ForeColor
    $RichTextBox.BackColor       = $BackColor
    $RichTextBox.TabIndex        = $TabIndex

    $Form.Controls.Add($RichTextBox)

    New-Variable                 `
        -Name $Name              `
        -Value $RichTextBox      `
        -Scope Script
}

Function Create-Label {
        param(
        $Text        = "",
        $Name        = "",
        $PosX        = 0,
        $PosY        = 0,
        $Width       = 62,
        $Height      = 62,
        $Enabled     = $true,
        $ForeColor   = $SCRIPT:DEFAULT_COLOR_FOREGROUND,
        $BackColor   = $SCRIPT:DEFAULT_COLOR_BACKGROUND,
        $TabIndex
    )

    # ===========================================================================================
    $Label                            = New-Object System.Windows.Forms.Label
    $RowStyle                         = New-Object System.Windows.Forms.RowStyle
    $ColumnStyle                      = New-Object System.Windows.Forms.ColumnStyle
    $TableLayoutPanel                 = New-Object System.Windows.Forms.TableLayoutPanel
    # ===========================================================================================

    $Label.Font                       = $SCRIPT:DEFAULT_FONT
    $Label.Text                       = $Text
    $Label.Name                       = $Name
    $Label.Anchor                     = [System.Windows.Forms.AnchorStyles]::Left, [System.Windows.Forms.AnchorStyles]::Right
    $Label.Width                      = 50
    $Label.Height                     = 28
    $Label.Enabled                    = $Enabled
    $Label.AutoSize                   = $false
    $Label.ForeColor                  = $ForeColor
    $Label.BackColor                  = $SCRIPT:DEFAULT_COLOR_BACKGROUND
    $Label.TextAlign                  = 'MiddleCenter'
    $Label.BorderStyle                = [System.Windows.Forms.BorderStyle]::None
    $Label.TabIndex                   = $TabIndex

    $TableLayoutPanel.Name            = "tlp_$Name"
    $TableLayoutPanel.Width           = $Width
    $TableLayoutPanel.Height          = $Height
    $TableLayoutPanel.AutoSize        = $false
    $TableLayoutPanel.Location        = New-Object System.Drawing.Point($PosX, $PosY)
    $TableLayoutPanel.RowCount        = 1
    $TableLayoutPanel.ColumnCount     = 1
    $TableLayoutPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single
    $TableLayoutPanel.BackColor       = $SCRIPT:DEFAULT_COLOR_BACKGROUND

    [void]$TableLayoutPanel.RowStyles.Add([System.Windows.Forms.RowStyle]::new("Percent", "50"))
    [void]$TableLayoutPanel.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new("Percent", "50"))

    $TableLayoutPanel.Controls.Add($Label)

    New-Variable                 `
        -Name $Name              `
        -Value $Label            `
        -Scope Script
    New-Variable                 `
        -Name "tlp_$Name"        `
        -Value $TableLayoutPanel `
        -Scope Script
    
    # ===========================================================================================
    $Form.Controls.Add($TableLayoutPanel)
    # ===========================================================================================
}

Function Clear-RichTextBoxes {
    $RTB1.Text = ""
    $RTB2.Text = ""
    $RTB3.Text = ""
}

Function Add-WordsToRichTextBoxes($Word) {
    $RTB1_Rows = ($RTB1.Text -split "`r`n").Length
    $RTB2_Rows = ($RTB2.Text -split "`r`n").Length
    $RTB3_Rows = ($RTB3.Text -split "`r`n").Length

    if ($RTB1_Rows -lt 18) {
        $RTB1.Text += "• $Word`r`n"
    } elseif ($RTB2_Rows -lt 18) {
        $RTB2.Text += "• $Word`r`n"
    } elseif ($RTB3_Rows -lt 18) {
        $RTB3.Text += "• $Word`r`n"
    }
}

Function ConvertFrom-Base64 {
        param(
            [Parameter(ValueFromPipeline = $true)]
            [String]$String
        )
        [System.Text.Encoding]::Utf8.GetString([System.Convert]::FromBase64String($String)) 
    }

Function Calculate-PossibleWords {
    [String]$Column1InvalidLetters = ''
    [String]$Column2InvalidLetters = ''
    [String]$Column3InvalidLetters = ''
    [String]$Column4InvalidLetters = ''
    [String]$Column5InvalidLetters = ''
    [String]$LettersInWord = ''
    [String]$Column1 = ''
    [String]$Column2 = ''
    [String]$Column3 = ''
    [String]$Column4 = ''
    [String]$Column5 = ''

    $LettersNotInWord = $SCRIPT:CONTROLS | ? {[String]$_.GetType() -eq 'System.Windows.Forms.Label'} | ? {$_.BackColor -eq $SCRIPT:DEFAULT_COLOR_NOT_IN_WORD}

    if ($LettersNotInWord.Length -eq 0) { return }

    $LettersNotInWord | % {
        $Column1InvalidLetters += $_.Text
        $Column2InvalidLetters += $_.Text
        $Column3InvalidLetters += $_.Text
        $Column4InvalidLetters += $_.Text
        $Column5InvalidLetters += $_.Text
    }

    $Column1InvalidLetters

    $LettersInWrongSpot = $SCRIPT:CONTROLS | ? {[String]$_.GetType() -eq 'System.Windows.Forms.Label'} | ? {$_.BackColor -eq $SCRIPT:DEFAULT_COLOR_WRONG_SPOT}
    $LettersInWrongSpot | % {
        $Name = $_.Name
        $Text = $_.Text
        Switch -Regex ($Name) {
            "C1$" { $Column1InvalidLetters += $Text }
            "C2$" { $Column2InvalidLetters += $Text }
            "C3$" { $Column3InvalidLetters += $Text }
            "C4$" { $Column4InvalidLetters += $Text }
            "C5$" { $Column5InvalidLetters += $Text }
        }
    }

    $LettersInCorrectSpot = $SCRIPT:CONTROLS | ? {[String]$_.GetType() -eq 'System.Windows.Forms.Label'} | ? {$_.BackColor -eq $SCRIPT:DEFAULT_COLOR_CORRECT_SPOT}
    $LettersInCorrectSpot | % {
        $Name = $_.Name
        $Text = $_.Text
        Switch -Regex ($Name) {
            "C1$" { $Column1 = $Text }
            "C2$" { $Column2 = $Text }
            "C3$" { $Column3 = $Text }
            "C4$" { $Column4 = $Text }
            "C5$" { $Column5 = $Text }
        }
    }

    if ($Column1) { $R_C1 = $Column1 } else { $R_C1 = "[^$Column1InvalidLetters]" }
    if ($Column2) { $R_C2 = $Column2 } else { $R_C2 = "[^$Column2InvalidLetters]" }
    if ($Column3) { $R_C3 = $Column3 } else { $R_C3 = "[^$Column3InvalidLetters]" }
    if ($Column4) { $R_C4 = $Column4 } else { $R_C4 = "[^$Column4InvalidLetters]" }
    if ($Column5) { $R_C5 = $Column5 } else { $R_C5 = "[^$Column5InvalidLetters]" }
    
    $Regex = "$R_C1$R_C2$R_C3$R_C4$R_C5"

    $LettersInWrongSpot | % {
        [String]$LettersInWord += $_.Text
    }

    Switch ($LettersInWord.ToCharArray().Length) {
        1 {
            $indexA = $LettersInWord.ToCharArray()[0]
            $PossibleWords = $SCRIPT:WORDLIST | ? { $_.Length -eq 5} | ? {$_ -match "$Regex"} | ? {$_ -match $indexA}
        }
        2 {
            $indexA = $LettersInWord.ToCharArray()[0]
            $indexB = $LettersInWord.ToCharArray()[1]
            $PossibleWords = $SCRIPT:WORDLIST | ? { $_.Length -eq 5} | ? {$_ -match "$Regex"} | ? {$_ -match $indexA} | ? {$_ -match $indexB}
        }
        3 {
            $indexA = $LettersInWord.ToCharArray()[0]
            $indexB = $LettersInWord.ToCharArray()[1]
            $indexC = $LettersInWord.ToCharArray()[2]
            $PossibleWords = $SCRIPT:WORDLIST | ? { $_.Length -eq 5} | ? {$_ -match "$Regex"} | ? {$_ -match $indexA} | ? {$_ -match $indexB} | ? {$_ -match $indexC}
        }
        4 {
            $indexA = $LettersInWord.ToCharArray()[0]
            $indexB = $LettersInWord.ToCharArray()[1]
            $indexC = $LettersInWord.ToCharArray()[2]
            $indexD = $LettersInWord.ToCharArray()[3]
            $PossibleWords = $SCRIPT:WORDLIST | ? { $_.Length -eq 5} | ? {$_ -match "$Regex"} | ? {$_ -match $indexA} | ? {$_ -match $indexB} | ? {$_ -match $indexC} | ? {$_ -match $indexD}
        }
        5 {
            $indexA = $LettersInWord.ToCharArray()[0]
            $indexB = $LettersInWord.ToCharArray()[1]
            $indexC = $LettersInWord.ToCharArray()[2]
            $indexD = $LettersInWord.ToCharArray()[3]
            $indexE = $LettersInWord.ToCharArray()[4]
            $PossibleWords = $SCRIPT:WORDLIST | ? { $_.Length -eq 5} | ? {$_ -match "$Regex"} | ? {$_ -match $indexA} | ? {$_ -match $indexB} | ? {$_ -match $indexC} | ? {$_ -match $indexD} | ? {$_ -match $indexE}
        }
    }

    Clear-RichTextBoxes
    foreach ($Word in $PossibleWords) {
        Add-WordsToRichTextBoxes -Word $Word
    }
}

# ==========================================================================
$Form                  = New-Object System.Windows.Forms.Form
# ==========================================================================
$Form.Text             = "Wordle Solver"
$Form.ForeColor        = $SCRIPT:DEFAULT_COLOR_FOREGROUND
$Form.BackColor        = $SCRIPT:DEFAULT_COLOR_BACKGROUND
$Form.ClientSize       = "745, 418"
$Form.FormBorderStyle  = [System.Windows.Forms.FormBorderStyle]::FixedSingle
# ==========================================================================

# =================== #

Create-Label                                                           `
    -Name       lbl_R1C1                                               `
    -PosX       10                                                     `
    -PosY       10                                                     `
    -TabIndex   0                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R1C2                                               `
    -PosX       $($tlp_lbl_R1C1.Location.X + $tlp_lbl_R1C1.Width + 5)  `
    -PosY       $($tlp_lbl_R1C1.Location.Y)                            `
    -TabIndex   1                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R1C3                                               `
    -PosX       $($tlp_lbl_R1C2.Location.X + $tlp_lbl_R1C2.Width + 5)  `
    -PosY       $($tlp_lbl_R1C1.Location.Y)                            `
    -TabIndex   2                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R1C4                                               `
    -PosX       $($tlp_lbl_R1C3.Location.X + $tlp_lbl_R1C3.Width + 5)  `
    -PosY       $($tlp_lbl_R1C1.Location.Y)                            `
    -TabIndex   3                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R1C5                                               `
    -PosX       $($tlp_lbl_R1C4.Location.X + $tlp_lbl_R1C4.Width + 5)  `
    -PosY       $($tlp_lbl_R1C1.Location.Y)                            `
    -TabIndex   4                                                      `
    -Enabled    $true

# =================== #

Create-Label                                                           `
    -Name       lbl_R2C1                                               `
    -PosX       10                                                     `
    -PosY       $($tlp_lbl_R1C1.Location.Y + $tlp_lbl_R1C1.Height + 5) `
    -TabIndex   5                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R2C2                                               `
    -PosX       $($tlp_lbl_R2C1.Location.X + $tlp_lbl_R2C1.Width + 5)  `
    -PosY       $($tlp_lbl_R2C1.Location.Y)                            `
    -TabIndex   6                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R2C3                                               `
    -PosX       $($tlp_lbl_R2C2.Location.X + $tlp_lbl_R2C2.Width + 5)  `
    -PosY       $($tlp_lbl_R2C1.Location.Y)                            `
    -TabIndex   7                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R2C4                                                `
    -PosX       $($tlp_lbl_R2C3.Location.X + $tlp_lbl_R2C3.Width + 5)    `
    -PosY       $($tlp_lbl_R2C1.Location.Y)                             `
    -TabIndex   8                                                      `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R2C5                                               `
    -PosX       $($tlp_lbl_R2C4.Location.X + $tlp_lbl_R2C4.Width + 5)  `
    -PosY       $($tlp_lbl_R2C1.Location.Y)                            `
    -TabIndex   9                                                      `
    -Enabled    $true

# =================== #

Create-Label                                                           `
    -Name       lbl_R3C1                                               `
    -PosX       10                                                     `
    -PosY       $($tlp_lbl_R2C1.Location.Y + $tlp_lbl_R2C1.Height + 5) `
    -TabIndex   10                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R3C2                                               `
    -PosX       $($tlp_lbl_R3C1.Location.X + $tlp_lbl_R3C1.Width + 5)  `
    -PosY       $($tlp_lbl_R3C1.Location.Y)                            `
    -TabIndex   11                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R3C3                                               `
    -PosX       $($tlp_lbl_R3C2.Location.X + $tlp_lbl_R3C2.Width + 5)  `
    -PosY       $($tlp_lbl_R3C1.Location.Y)                            `
    -TabIndex   12                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R3C4                                               `
    -PosX       $($tlp_lbl_R3C3.Location.X + $tlp_lbl_R3C3.Width + 5)  `
    -PosY       $($tlp_lbl_R3C1.Location.Y)                            `
    -TabIndex   13                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R3C5                                               `
    -PosX       $($tlp_lbl_R3C4.Location.X + $tlp_lbl_R3C4.Width + 5)  `
    -PosY       $($tlp_lbl_R3C1.Location.Y)                            `
    -TabIndex   14                                                     `
    -Enabled    $true

# =================== #

Create-Label                                                           `
    -Name       lbl_R4C1                                               `
    -PosX       10                                                     `
    -PosY       $($tlp_lbl_R3C1.Location.Y + $tlp_lbl_R3C1.Height + 5) `
    -TabIndex   15                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R4C2                                               `
    -PosX       $($tlp_lbl_R4C1.Location.X + $tlp_lbl_R4C1.Width + 5)  `
    -PosY       $($tlp_lbl_R4C1.Location.Y)                            `
    -TabIndex   16                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R4C3                                               `
    -PosX       $($tlp_lbl_R4C2.Location.X + $tlp_lbl_R4C2.Width + 5)  `
    -PosY       $($tlp_lbl_R4C1.Location.Y)                            `
    -TabIndex   17                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R4C4                                               `
    -PosX       $($tlp_lbl_R4C3.Location.X + $tlp_lbl_R4C3.Width + 5)  `
    -PosY       $($tlp_lbl_R4C1.Location.Y)                            `
    -TabIndex   18                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R4C5                                               `
    -PosX       $($tlp_lbl_R4C4.Location.X + $tlp_lbl_R4C4.Width + 5)  `
    -PosY       $($tlp_lbl_R4C1.Location.Y)                            `
    -TabIndex   19                                                     `
    -Enabled    $true

# =================== #

Create-Label                                                           `
    -Name       lbl_R5C1                                               `
    -PosX       10                                                     `
    -PosY       $($tlp_lbl_R4C1.Location.Y + $tlp_lbl_R4C1.Height + 5) `
    -TabIndex   20                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R5C2                                               `
    -PosX       $($tlp_lbl_R5C1.Location.X + $tlp_lbl_R5C1.Width + 5)  `
    -PosY       $($tlp_lbl_R5C1.Location.Y)                            `
    -TabIndex   21                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R5C3                                               `
    -PosX       $($tlp_lbl_R5C2.Location.X + $tlp_lbl_R5C2.Width + 5)  `
    -PosY       $($tlp_lbl_R5C1.Location.Y)                            `
    -TabIndex   22                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R5C4                                               `
    -PosX       $($tlp_lbl_R5C3.Location.X + $tlp_lbl_R5C3.Width + 5)  `
    -PosY       $($tlp_lbl_R5C1.Location.Y)                            `
    -TabIndex   23                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R5C5                                               `
    -PosX       $($tlp_lbl_R5C4.Location.X + $tlp_lbl_R5C4.Width + 5)  `
    -PosY       $($tlp_lbl_R5C1.Location.Y)                            `
    -TabIndex   24                                                     `
    -Enabled    $true

# =================== #

Create-Label                                                           `
    -Name       lbl_R6C1                                               `
    -PosX       10                                                     `
    -PosY       $($tlp_lbl_R5C1.Location.Y + $tlp_lbl_R5C1.Height + 5) `
    -TabIndex   25                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R6C2                                               `
    -PosX       $($tlp_lbl_R6C1.Location.X + $tlp_lbl_R6C1.Width + 5)  `
    -PosY       $($tlp_lbl_R6C1.Location.Y)                            `
    -TabIndex   26                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R6C3                                               `
    -PosX       $($tlp_lbl_R6C2.Location.X + $tlp_lbl_R6C2.Width + 5)  `
    -PosY       $($tlp_lbl_R6C1.Location.Y)                            `
    -TabIndex   27                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R6C4                                               `
    -PosX       $($tlp_lbl_R6C3.Location.X + $tlp_lbl_R6C3.Width + 5)  `
    -PosY       $($tlp_lbl_R6C1.Location.Y)                            `
    -TabIndex   28                                                     `
    -Enabled    $true

Create-Label                                                           `
    -Name       lbl_R6C5                                               `
    -PosX       $($tlp_lbl_R6C4.Location.X + $tlp_lbl_R6C4.Width + 5)  `
    -PosY       $($tlp_lbl_R6C1.Location.Y)                            `
    -TabIndex   29                                                     `
    -Enabled    $true

Create-RichTextBox                                                     `
    -Name       rtb1                                                   `
    -PosX       $($tlp_lbl_R1C5.Location.X + $tlp_lbl_R1C5.Width + 10) `
    -PosY       10                                                     `
    -Width      125                                                    `
    -Height     397                                                    `
    -TabIndex   30                                                     `
    -Enabled    $true                                                  `
    -ForeColor  $SCRIPT:DEFAULT_COLOR_FOREGROUND                       `
    -BackColor  $SCRIPT:DEFAULT_COLOR_BACKGROUND

Create-RichTextBox                                                     `
    -Name       rtb2                                                   `
    -PosX       $($rtb1.Location.X + $rtb1.Width + 5)                 `
    -PosY       10                                                     `
    -Width      125                                                    `
    -Height     397                                                    `
    -TabIndex   30                                                     `
    -Enabled    $true                                                  `
    -ForeColor  $SCRIPT:DEFAULT_COLOR_FOREGROUND                       `
    -BackColor  $SCRIPT:DEFAULT_COLOR_BACKGROUND

Create-RichTextBox                                                     `
    -Name       rtb3                                                   `
    -PosX       $($rtb2.Location.X + $rtb2.Width + 5)                 `
    -PosY       10                                                     `
    -Width      125                                                    `
    -Height     397                                                    `
    -TabIndex   30                                                     `
    -Enabled    $true                                                  `
    -ForeColor  $SCRIPT:DEFAULT_COLOR_FOREGROUND                       `
    -BackColor  $SCRIPT:DEFAULT_COLOR_BACKGROUND

# =================== #

Function SetCellColor($ColorVariable) {
    Switch ($this.GetType()) {
        System.Windows.Forms.TableLayoutPanel {
            $this.BackColor = $ColorVariable
            $this.Controls[0].BackColor = $ColorVariable
        }
        System.Windows.Forms.Label {
            $this.BackColor = $ColorVariable
            $this.Parent.BackColor = $ColorVariable
        }
    }
}

$lbl_KeyPressed = {
    Function GetNextTile {      
        $Labels = $SCRIPT:CONTROLS | ? {$_.Name -match "^lbl_R"}
        $FirstLabel = $Labels | Sort-Object -Property TabIndex | Select -First 1
        $LastLabel = $Labels | Sort-Object -Property TabIndex | Select -Last 1

        foreach ($Label in ($Labels | Sort-Object -Property TabIndex)) {
            $HasText = [Boolean]($Label.Text.Length -eq 1)

            if (!$HasText) {
                $NextLabel = $Label
                return $NextLabel
            }
        }
        return $LastLabel
    }

    Function GetPreviousTile {
        $Labels = $SCRIPT:CONTROLS | ? {$_.Name -match "^lbl_R"}
        $NextLabel = $Labels | Sort-Object -Property TabIndex | Select -First 1

        foreach ($Label in ($Labels | Sort-Object -Property TabIndex)) {
            $HasText = [Boolean]($Label.Text.Length -eq 1)

            if (!$HasText) {
                $NextLabel = $Label
                $TabIndex = $NextLabel.TabIndex
                return $Labels | ? {$_.TabIndex -eq ($TabIndex - 1)}
            }
        }
        return $Labels | Sort-Object -Property TabIndex | Select -Last 1
    }

    $KeyCode = $args[1].KeyCode
    if ($KeyCode -match "^[A-Z]$") {
        $NextTile = GetNextTile
        if ($NextTile) {
            $NextTile.Text = $KeyCode.ToString().ToUpper()
            if ($NextTile.BackColor -eq $SCRIPT:DEFAULT_COLOR_BACKGROUND) {
                $NextTile.BackColor = $SCRIPT:DEFAULT_COLOR_NOT_IN_WORD
                $NextTile.Parent.BackColor = $SCRIPT:DEFAULT_COLOR_NOT_IN_WORD
            }
        }
    } elseif ($KeyCode -eq [System.Windows.Forms.Keys]::Back) {
        $PreviousTile = GetPreviousTile
        if ($PreviousTile) {
            $PreviousTile.Text = ""
            if ($PreviousTile.BackColor -ne $SCRIPT:DEFAULT_COLOR_BACKGROUND) {
                $PreviousTile.BackColor = $SCRIPT:DEFAULT_COLOR_BACKGROUND
                $PreviousTile.Parent.BackColor = $SCRIPT:DEFAULT_COLOR_BACKGROUND
            }
        }
    } elseif ($KeyCode -eq [System.Windows.Forms.Keys]::Return) {
        Calculate-PossibleWords
    }
}

Enumerate-Controls $Form

$SCRIPT:CONTROLS | ? {$_.Name -match "^tlp_lbl_R"} | % {
    $_.Add_Click({SetCellColor -ColorVariable $SCRIPT:DEFAULT_COLOR_WRONG_SPOT})
    $_.Add_DoubleClick({SetCellColor -ColorVariable $SCRIPT:DEFAULT_COLOR_CORRECT_SPOT})
    $_.Add_MouseUp({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            SetCellColor -ColorVariable $SCRIPT:DEFAULT_COLOR_NOT_IN_WORD
        }
    })
}

$SCRIPT:CONTROLS | ? {$_.Name -match "^lbl_R"} | % {
    $_.Add_Click({SetCellColor -ColorVariable $SCRIPT:DEFAULT_COLOR_WRONG_SPOT})
    $_.Add_DoubleClick({SetCellColor -ColorVariable $SCRIPT:DEFAULT_COLOR_CORRECT_SPOT})
    $_.Add_MouseUp({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            SetCellColor -ColorVariable $SCRIPT:DEFAULT_COLOR_NOT_IN_WORD
        }
    })
}

$Form.Add_KeyUp($lbl_KeyPressed)


[void]$Form.ShowDialog()
