# Create a Window
$window = New-Object System.Windows.Window
$window.Title = "Microsoft Rewards Additional Thumbnails Keywords"
$window.width = 680  # Set a wider window width
$window.SizeToContent = "height"  # Automatically adjust width and height based on content
$window.WindowStartupLocation = "CenterScreen"
$window.Background = [System.Windows.Media.Brushes]::Black

# Create a StackPanel for vertical layout
$stackPanel = New-Object System.Windows.Controls.StackPanel
$stackPanel.HorizontalAlignment = "Center"  # Center the stack content horizontally
$window.Content = $stackPanel

$toClick = [System.Collections.Generic.List[string]]::new()

# Function to add keyword if not already in the list and update the label
function AddKeyword {
    param ($keyword)

    if (-not $toClick.Contains($keyword)) {
        $toClick.Add($keyword)
        # Update the label with the updated list of keywords
        $selectedThumbnailsText.Inlines.Clear()
        
        # Add underlined "Vignettes choisies :"
        $underlineRun = New-Object System.Windows.Documents.Run("Vignettes choisies : ")
        $underlineRun.TextDecorations = [System.Windows.TextDecorations]::Underline
        $selectedThumbnailsText.Inlines.Add($underlineRun)
        
        # Add the list of keywords
        $selectedThumbnailsText.Inlines.Add((New-Object System.Windows.Documents.Run(($toClick -join ", "))))
    }
}



# Function to remove the last entered keyword if the list is not empty
function RemoveLastKeyword {
    if ($toClick.Count -gt 0) {
        $toClick.RemoveAt($toClick.Count - 1)
        UpdateSelectedThumbnails
    }
}

function UpdateSelectedThumbnails {
    $selectedThumbnailsText.Inlines.Clear()
    
    # Add underlined "Vignettes choisies :"
    $underlineRun = New-Object System.Windows.Documents.Run("Vignettes choisies : ")
    $underlineRun.TextDecorations = [System.Windows.TextDecorations]::Underline
    $selectedThumbnailsText.Inlines.Add($underlineRun)
    
    # Add the list of keywords
    if ($toClick.Count -gt 0) {
        $selectedThumbnailsText.Inlines.Add((New-Object System.Windows.Documents.Run(($toClick -join ", "))))
    }
}



# Function to add all keywords
function AddAllKeywords {
    $allKeywords = @('convertion', 'symptomes', 'cuisiner', 'shop', 'hotel', 'maisons', 'meteo', 'win', 'economie', 'itineraire', 'jeu', 'chanson', 'traduction', 'film', 'heure', 'vocabulaire', 'elections', 'postes', 'vol')
    foreach ($keyword in $allKeywords) {
        AddKeyword $keyword
    }
}

# Create functions for each button to call AddKeyword with the specific keyword
function FunctionOne { AddKeyword 'convertion' }
function FunctionTwo { AddKeyword 'symptomes' }
function FunctionThree { AddKeyword 'cuisiner' }
function FunctionFour { AddKeyword 'shop' }
function FunctionFive { AddKeyword 'hotel' }
function FunctionSix { AddKeyword 'maisons' }
function FunctionSeven { AddKeyword 'meteo' }
function FunctionEight { AddKeyword 'win' }
function FunctionNine { AddKeyword 'economie' }
function FunctionTen { AddKeyword 'itineraire' }
function FunctionEleven { Addkeyword 'jeu' }
function FunctionTwelve { Addkeyword 'chanson' }
function FunctionThirteen { Addkeyword 'traduction' }
function FunctionFourteen { Addkeyword 'film' }
function FunctionFifteen { Addkeyword 'heure' }
function FunctionSixteen { Addkeyword 'vocabulaire' }
function FunctionSeventeen { Addkeyword 'elections' }
function FunctionEighteen { Addkeyword 'postes' }
function FunctionNineteen { Addkeyword 'vol' }

# Create a Grid with two columns
$grid = New-Object System.Windows.Controls.Grid
$grid.HorizontalAlignment = "Center"
$grid.Margin = "10"

# Define two columns for the Grid
$col1 = New-Object System.Windows.Controls.ColumnDefinition
$col2 = New-Object System.Windows.Controls.ColumnDefinition
$grid.ColumnDefinitions.Add($col1)
$grid.ColumnDefinitions.Add($col2)

$window.Content = $grid

# Add Subtitle
$Main = New-Object System.Windows.Controls.Label
$Main.Content = "Choisissez les vignettes additionnelles : "
$Main.FontSize = 30
$Main.HorizontalAlignment = "Center"
$Main.Margin = "5,0,5,0"
$Main.Foreground = [System.Windows.Media.Brushes]::Green
$grid.Children.Add($Main)
[System.Windows.Controls.Grid]::SetColumnSpan($Main, 2)  # Centered across both columns

# Create Buttons and assign their Click events
$buttonData = @(
    @{ Content = "Conversion rapide de monnaie"; Action = { FunctionOne } },
    @{ Content = "Vous avez des symptômes?"; Action = { FunctionTwo } },
    @{ Content = "Trop fatigué pour cuisiner ce soir?"; Action = { FunctionThree } },
    @{ Content = "Faites vos achats plus vite"; Action = { FunctionFour } },
    @{ Content = "Trouvez des emplacements pour rester!"; Action = { FunctionFive } },
    @{ Content = "Maisons près de chez vous!"; Action = { FunctionSix } },
    @{ Content = "Vérifier la météo"; Action = {FunctionSeven}},
    @{ Content = "Qui a gagné ?"; Action = {FunctionEight}},
    @{ Content = "Comment se porte l'économie ?"; Action = {FunctionNine}},
    @{ Content = "Trouver un endroit à découvrir"; Action = {FunctionTen}},
    @{ Content = "Temps de jeu"; Action = {FunctionEleven}},
    @{ Content = "Rechercher paroles de chanson"; Action = {FunctionTwelve}},
    @{ Content = "Traduisez tout !"; Action = {FunctionThirteen}},
    @{ Content = "Et si nous regardions ce film une nouvelle fois ?"; Action = {FunctionFourteen}},
    @{ Content = "Quelle heure est-il ?"; Action = {FunctionFifteen}},
    @{ Content = "Enrichissez votre vocabulaire"; Action = {FunctionSixteen}},
    @{ Content = "Suivez les élections"; Action = {FunctionSeventeen}},
    @{ Content = "Consulter postes à pourvoir"; Action = {FunctionEighteen}},
    @{ Content = "Planifiez une petite escapade"; Action = {FunctionNineteen}}
)







# Track the row position
$row = 1
for ($i = 0; $i -lt $buttonData.Length; $i++) {
    $button = New-Object System.Windows.Controls.Button
    $button.Content = $buttonData[$i].Content
    $button.Width = 300  # Adjust width for two columns
    $button.Margin = "5,5,5,5"
    $button.Background = [System.Windows.Media.Brushes]::Transparent
    $button.Foreground = [System.Windows.Media.Brushes]::Green
    $button.FontSize = 20
    $button.Add_Click($buttonData[$i].Action)

    # Determine column and set row position
    $column = $i % 2
    $grid.Children.Add($button)
    $button.SetValue([System.Windows.Controls.Grid]::RowProperty, $row)
    $button.SetValue([System.Windows.Controls.Grid]::ColumnProperty, $column)

    # Move to next row after adding two buttons
    if ($column -eq 1) { $row++ }
}

# Add rows as needed based on button count
for ($i = 1; $i -le ($row+8); $i++) {
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
}












# Add "Add All" and "Remove Last" buttons below the grid
$addAllButton = New-Object System.Windows.Controls.Button
$addAllButton.Content = "Ajouter toutes les vignettes"
$addAllButton.Width = 400
$addAllButton.Margin = "5,15,5,0"
$addAllButton.Background = [System.Windows.Media.Brushes]::Transparent
$addAllButton.Foreground = [System.Windows.Media.Brushes]::Green
$addAllButton.FontSize = 20
$addAllButton.FontWeight = 'bold'
$addAllButton.Add_Click({ AddAllKeywords })
$grid.Children.Add($addAllButton)
$addAllButton.SetValue([System.Windows.Controls.Grid]::RowProperty, $row +1)
[System.Windows.Controls.Grid]::SetColumnSpan($addAllButton, 2)

$removeLastButton = New-Object System.Windows.Controls.Button
$removeLastButton.Content = "Supprimer le dernier mot-clé"
$removeLastButton.Width = 400
$removeLastButton.Margin = "5,5,5,0"
$removeLastButton.Background = [System.Windows.Media.Brushes]::Transparent
$removeLastButton.Foreground = [System.Windows.Media.Brushes]::Green
$removeLastButton.FontSize = 20
$removeLastButton.FontWeight = 'bold'
$removeLastButton.Add_Click({ RemoveLastKeyword })
$grid.Children.Add($removeLastButton)
$removeLastButton.SetValue([System.Windows.Controls.Grid]::RowProperty, $row +2)
[System.Windows.Controls.Grid]::SetColumnSpan($removeLastButton, 2)

# TextBlock to show selected thumbnails with "Vignettes choisies" underlined
$selectedThumbnailsText = New-Object System.Windows.Controls.TextBlock
$selectedThumbnailsText.HorizontalAlignment = "Center"
$selectedThumbnailsText.Margin = "15,0,15,20"
$selectedThumbnailsText.FontSize = 20
$selectedThumbnailsText.TextWrapping = [System.Windows.TextWrapping]::Wrap  # Enable text wrapping
$selectedThumbnailsText.Foreground = [System.Windows.Media.Brushes]::Green
$grid.Children.Add($selectedThumbnailsText)
$selectedThumbnailsText.SetValue([System.Windows.Controls.Grid]::RowProperty, $row +3)
[System.Windows.Controls.Grid]::SetColumnSpan($selectedThumbnailsText, 2)
[System.Windows.Controls.Grid]::setRowSpan($selectedThumbnailsText, 3)


# Initialize "Vignettes choisies" with underline
$underlineRun = New-Object System.Windows.Documents.Run("Vignettes choisies : ")
$underlineRun.TextDecorations = [System.Windows.TextDecorations]::Underline
$selectedThumbnailsText.Inlines.Add($underlineRun)



# Ajouter un bouton pour lancer les recherches en fonction des mots-clés sélectionnés
$launchSearchButton = New-Object System.Windows.Controls.Button
$launchSearchButton.Content = "Lancer les recherches"
$launchSearchButton.Width = 400
$launchSearchButton.Margin = "5,0,5,10"
$launchSearchButton.Background = [System.Windows.Media.Brushes]::Transparent
$launchSearchButton.Foreground = [System.Windows.Media.Brushes]::White
$launchSearchButton.FontSize = 20
$launchSearchButton.FontWeight = 'bold'
$launchSearchButton.Add_Click({ LaunchSearch })
$grid.Children.Add($launchSearchButton)
$launchSearchButton.SetValue([System.Windows.Controls.Grid]::RowProperty, $row +6)
[System.Windows.Controls.Grid]::SetColumnSpan($launchSearchButton, 2)













# Fonction pour ouvrir la recherche Bing en fonction des mots-clés sélectionnés
function LaunchSearch {
    foreach ($keyword in $toClick) {
        switch ($keyword) {
            'convertion' {
                # Recherche de conversion USD à EURO directement sur Bing
                Start-Process "https://www.bing.com/search?q=convertion+usd+euro"
                Start-Sleep -Seconds 5
            }
            'symptomes' {
                # Recherche de symptômes de grippe
                Start-Process "https://www.bing.com/search?q=symptomes+grippe"
                Start-Sleep -Seconds 5
            }
            'cuisiner' {
                # Recherche pour un restaurant Buffalo Grill, avec département à préciser
                $department = "charente" # Remplacez par le département souhaité
                Start-Process "https://www.bing.com/search?q=buffalo+grill+$department"
                Start-Sleep -Seconds 5
            }
            'shop' {
                # Lien direct vers Bing Shopping
                Start-Process "https://www.bing.com/search?q=prix+ordinateur+portable&qs=SC&pq=prix+ordinatzeur&sc=10-16&cvid=5F8A3D8B45C042458174CBA32417B8F2&FORM=QBRE&sp=1&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
            'hotel' {
                # Recherche de location de chambres en anglais sur bing.com (US)
                Start-Process "https://www.bing.com/search?q=Rent+room+hotel&cc=us"
                Start-Sleep -Seconds 5
            }
            'maisons' {
                # Recherche immobilière dans une ville spécifique, à préciser
                $city = "Angouleme" # Remplacez par votre ville
                Start-Process "https://www.bing.com/search?q=real+estate+$city&cc=us"
                Start-Sleep -Seconds 5
            }
            'meteo' {
                $city = "Angouleme" # Remplacez par votre ville
                Start-Process "https://www.bing.com/search?q=meteo+angouleme&form=QBLH&sp=-1&ghc=1&lq=0&pq=meteo+angouleme&sc=10-15&qs=n&sk=&cvid=F23A6E6B70A94DF3896E99524492AF2A&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            'win' {
                Start-Process "https://www.bing.com/?cc=us"
                Start-Process "https://www.bing.com/search?q=score+miami+heat&qs=MT&pq=score+&sk=MT2LT1&sc=10-6&cvid=E12C36414DF54F03A96D70371EE10D04&FORM=QBLH&sp=4&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
            'economie' {
                Start-Process "https://www.bing.com/?cc=us"
                Start-Process "https://www.bing.com/search?q=action+amazon&qs=n&form=QBRE&sp=-1&ghc=1&lq=0&pq=action+amazon&sc=10-13&sk=&cvid=E6492F03B82A49EA94D06CBDEA91FF65&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            'itineraire' {
                Start-Process "https://www.bing.com/search?q=itin%c3%a9raire+angoul%c3%aame+toulouse&qs=UT&pq=itin%c3%a9raire+angoul%c3%aame+&sc=10-21&cvid=A0733621207C4C0480535828AD68DDF3&FORM=QBRE&sp=1&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
            'jeu' {
                Start-Process "https://www.bing.com/?cc=us"
                Start-Process "https://www.bing.com/search?q=mario+bros+game&form=QBLH&sp=-1&ghc=1&lq=0&pq=mario+bros+game&sc=10-15&qs=n&sk=&cvid=F11810355D884613BF0874064EB570C4&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            'chanson' {
                Start-Process "https://www.bing.com/search?form=&q=lyrics+paint+it+black&form=QBLH&sp=-1&lq=0&pq=lyrics+paint+it+black&sc=0-21&qs=n&sk=&cvid=4CD63C7F651443999EB9BE6C637861BF&ghsh=0&ghacc=0&ghpl=0
"
                Start-Sleep -Seconds 5
            }
            'traduction' {
                Start-Process "https://www.bing.com/search?q=foret+en+espagnol&qs=n&form=QBRE&sp=-1&ghc=1&lq=0&pq=foret+en+espagnol&sc=3-17&sk=&cvid=07A5D7113A5C4048A995DD3E9738A41C&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            'film' {
                Start-Process "https://www.bing.com/search?q=titanic+film&qs=AS&pq=titanic+&sc=10-8&cvid=C02A3307DD2345AD82C03C6F41FC6010&FORM=QBLH&sp=1&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
            'heure' {
                Start-Process "https://www.bing.com/search?q=heure+new+york&qs=n&form=QBRE&sp=-1&ghc=2&lq=0&pq=heure+ne&sc=10-8&sk=&cvid=0C883D28B23B49C8B6550744D6FD81CE&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            'vocabulaire' {
                Start-Process "https://www.bing.com/search?q=pamphlet+def&qs=SC&pq=panphlet+&sk=SC1&sc=10-9&cvid=EA6C1C699E754334ACA531BD45AFACE0&FORM=QBLH&sp=2&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
            'elections' {
                Start-Process "https://www.bing.com/search?q=elections+usa+2024&form=QBLH&sp=-1&ghc=2&lq=0&pq=elections+&sc=10-10&qs=n&sk=&cvid=091D0ED3C50B4FFE902B428A9F800208&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            'postes' {
                Start-Process "https://www.bing.com/search?q=postes+%c3%a0+pourvoir&qs=FT&pq=postes+&sc=10-7&cvid=2D82020FCB114649BE9B732F2BF6F9F9&FORM=QBLH&sp=1&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
            
            'vol' {
                Start-Process "https://www.bing.com/search?q=vol+paris+new+york&form=QBLH&sp=-1&ghc=2&lq=0&pq=vol+paris+new+york&sc=10-18&qs=n&sk=&cvid=363E024D553F466DB6C2DF04F87589D5&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            
        }
    }
    Stop-Process -Name "msedge" -Force
}












# Show the Window
$window.ShowDialog()
