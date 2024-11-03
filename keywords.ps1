# Create a Window
$window = New-Object System.Windows.Window
$window.Title = "Microsoft Rewards Additional Thumbnails Keywords"
$window.Width = 700  # Set a wider window width
$window.Height = 500  # Set an appropriate height
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

# Function to add all keywords
function AddAllKeywords {
    $allKeywords = @('convertion', 'symptomes', 'cuisiner', 'shop', 'hotel', 'maisons')
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
    @{ Content = "Maisons près de chez vous!"; Action = { FunctionSix } }
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

# TextBlock to show selected thumbnails with "Vignettes choisies" underlined
$selectedThumbnailsText = New-Object System.Windows.Controls.TextBlock
$selectedThumbnailsText.HorizontalAlignment = "Center"
$selectedThumbnailsText.Margin = "5,25,5,10"
$selectedThumbnailsText.FontSize = 20
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
                Start-Process "https://www.bing.com/shop?FORM=Z9LHS4"
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
