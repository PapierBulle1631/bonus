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
    $allKeywords = @('convertion', 'symptomes', 'cuisiner', 'shop', 'hotel', 'maisons', 'meteo', 'win', 'economie', 'itineraire')
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

# Add Subtitle
$Main = New-Object System.Windows.Controls.Label
$Main.Content = "Choisissez les vignettes additionnelles : "
$Main.FontSize = 30
$Main.HorizontalAlignment = "Center"
$Main.Margin = "5,0,5,10"
$Main.Foreground = [System.Windows.Media.Brushes]::Green
$stackPanel.Children.Add($Main)

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
    @{ Content = "Trouver un endroit à découvrir"; Action = {FunctionTen}}
)

foreach ($data in $buttonData) {
    $button = New-Object System.Windows.Controls.Button
    $button.Content = $data.Content
    $button.Width = 400
    $button.Margin = "5,5,5,0"
    $button.Background = [System.Windows.Media.Brushes]::Transparent
    $button.Foreground = [System.Windows.Media.Brushes]::Green
    $button.FontSize = 20
    $button.Add_Click($data.Action)
    $stackPanel.Children.Add($button)
}

# Create "Add All" button
$addAllButton = New-Object System.Windows.Controls.Button
$addAllButton.Content = "Ajouter toutes les vignettes"
$addAllButton.Width = 400
$addAllButton.Margin = "5,22,5,10"
$addAllButton.Background = [System.Windows.Media.Brushes]::Transparent
$addAllButton.Foreground = [System.Windows.Media.Brushes]::Green
$addAllButton.FontSize = 20
$addAllButton.FontWeight = 'bold' 
$addAllButton.Add_Click({ AddAllKeywords })
$stackPanel.Children.Add($addAllButton)

# Create "Remove Last" button
$removeLastButton = New-Object System.Windows.Controls.Button
$removeLastButton.Content = "Supprimer le dernier mot-clé"
$removeLastButton.Width = 400
$removeLastButton.Margin = "5,5,5,10"
$removeLastButton.Background = [System.Windows.Media.Brushes]::Transparent
$removeLastButton.Foreground = [System.Windows.Media.Brushes]::Green
$removeLastButton.FontSize = 20
$removeLastButton.FontWeight = 'bold'
$removeLastButton.Add_Click({ RemoveLastKeyword })
$stackPanel.Children.Add($removeLastButton)


# TextBlock to show selected thumbnails with "Vignettes choisies" underlined
$selectedThumbnailsText = New-Object System.Windows.Controls.TextBlock
$selectedThumbnailsText.HorizontalAlignment = "Center"
$selectedThumbnailsText.Margin = "15,25,15,10"
$selectedThumbnailsText.FontSize = 20
$selectedThumbnailsText.TextWrapping = [System.Windows.TextWrapping]::Wrap  # Enable text wrapping
$selectedThumbnailsText.Foreground = [System.Windows.Media.Brushes]::Green
$stackPanel.Children.Add($selectedThumbnailsText)

# Initialize "Vignettes choisies" with underline
$underlineRun = New-Object System.Windows.Documents.Run("Vignettes choisies : ")
$underlineRun.TextDecorations = [System.Windows.TextDecorations]::Underline
$selectedThumbnailsText.Inlines.Add($underlineRun)













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
                Start-Process "https://www.bing.com/search?q=score+miami+heat&qs=MT&pq=score+&sk=MT2LT1&sc=10-6&cvid=E12C36414DF54F03A96D70371EE10D04&FORM=QBLH&sp=4&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
            'economie' {
                Start-Process "https://www.bing.com/search?q=action+amazon&qs=n&form=QBRE&sp=-1&ghc=1&lq=0&pq=action+amazon&sc=10-13&sk=&cvid=E6492F03B82A49EA94D06CBDEA91FF65&ghsh=0&ghacc=0&ghpl=0"
                Start-Sleep -Seconds 5
            }
            'itineraire' {
                Start-Process "https://www.bing.com/search?q=itin%c3%a9raire+angoul%c3%aame+toulouse&qs=UT&pq=itin%c3%a9raire+angoul%c3%aame+&sc=10-21&cvid=A0733621207C4C0480535828AD68DDF3&FORM=QBRE&sp=1&ghc=1&lq=0"
                Start-Sleep -Seconds 5
            }
        }
    }
    Stop-Process -Name "msedge" -Force
}

# Ajouter un bouton pour lancer les recherches en fonction des mots-clés sélectionnés
$launchSearchButton = New-Object System.Windows.Controls.Button
$launchSearchButton.Content = "Lancer les recherches"
$launchSearchButton.Width = 400
$launchSearchButton.Margin = "5,22,5,10"
$launchSearchButton.Background = [System.Windows.Media.Brushes]::Transparent
$launchSearchButton.Foreground = [System.Windows.Media.Brushes]::White
$launchSearchButton.FontSize = 20
$launchSearchButton.FontWeight = 'bold'
$launchSearchButton.Add_Click({ LaunchSearch })
$stackPanel.Children.Add($launchSearchButton)











# Show the Window
$window.ShowDialog()
