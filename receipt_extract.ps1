# If you use Albertsons, Safeway, etc., you can get your digital receipts emailed to you.  
# With this script, you can extract the HTML from the messages and create a spreadsheet of where your money goes.  
# Fun note: I joke with my spouse that the only reason I made this was to prove how much I spent on bags from the store.  
# The true story is:  
#   A) I was bored, and this was fun to me.  
#   B) I spend close to a thousand dollars at these stores a month, and I wanted to know why.  
# With this, I get a date and time stamp so I can extract what times or days I’m going.  
# I get the price of an item over time so I can make my own inflation statistics for my area.  
# I can see what items I’m purchasing most often (so I can buy in bulk).  
# A pipe dream is if I can automate the process and store the data into a web app, then I can know what I should have at home so I don’t buy something I don’t need.  
# An example would be adding expiration times to food descriptions so the app could, in theory, tell me if the chicken I bought is still good and display that in a table.  
# If the web app had a mobile companion or an Alexa tie-in, I could report when something was used up.  
# Think about those smart fridges where you have to scan every item to create an inventory, but what if it’s not stored in the fridge?  
# By consuming the receipts, you get everything.  
# I can also extract how frequently I buy something and analyze spending patterns over time. 
#
#
# What "I" did was create a gmail account just for this.
# I added it to outlook on my device.
# You update line28 (replace "Account name" but keep the quotes example "albertsons@gmail.com")
# I will output a access database to your MyDocuments folder by default to a file call Receipts.accdb - You can change the path in the variable just below.
# Script updated to use AccessDB vs CSV and added duplicate checks in the event you requested the same receipt twice.
#
#set the email address
$emailaccountname = "account.name@gmail.com"

#set the output location
# Path to the Access database
$accessDbPath = "$([Environment]::GetFolderPath('MyDocuments'))\Receipts.accdb"

# Access database connection string
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$accessDbPath;"

# Create Access database if it does not exist
if (-not (Test-Path $accessDbPath)) {
    Write-Host "Creating Access database..."
    $catalog = New-Object -ComObject ADOX.Catalog
    $catalog.Create($connectionString)
    
    $connection = New-Object -ComObject ADODB.Connection
    $connection.Open($connectionString)
    
    $createTableQuery = @"
    CREATE TABLE Receipts (
        ID AUTOINCREMENT PRIMARY KEY,
        Email TEXT(255),
        TransactionNumber TEXT(50),
        AuthorizationTime TEXT(50),
        AuthorizationDate TEXT(50),
        ReferenceNumber TEXT(50),
        AuthorizationCode TEXT(50),
        Amount TEXT(50),
        CardEndingIn TEXT(50),
        Calculated TEXT(50),
        SalesTax TEXT(50),
        TaxesAndFees TEXT(50),
        Subtotal TEXT(50),
        TotalSavings TEXT(50),
        TotalItems TEXT(50),
        ProductName TEXT(255),
        ProductPrice TEXT(50),
        ProductQuantity TEXT(50),
        ProductRegularPrice TEXT(50)
    );
"@

    $connection.Execute($createTableQuery) | Out-Null
    $connection.Close()
    Write-Host "Database and table created successfully."
}

# Create the Outlook application object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Get the Inbox folder for your Gmail account
$inbox = $namespace.Folders.Item($emailaccountname)
$gmailFolder = $inbox.Folders.Item("Inbox")
$messages = $gmailFolder.Items

# Ensure there are at least three emails
if ($messages.Count -ge 1) {
    $connection = New-Object -ComObject ADODB.Connection
    $connection.Open($connectionString)

    for ($i = 1; $i -le $messages.Count; $i++) {
        $message = $messages.Item($i)
        $subject = $message.Subject
        $receivedTime = $message.ReceivedTime
        $htmlBody = $message.HTMLBody

        # Extract data using regex patterns
        $transactionNumberPattern = 'Transaction Number\s*</td>\s*<td[^>]*>\s*([\d\s]+)\s*</td>'
        $transactionNumber = ([regex]::Match($htmlBody, $transactionNumberPattern)).Groups[1].Value -replace ' '

        $authorizationTimePattern = 'Authorization Time\s*</td>\s*<td[^>]*>\s*([\d:]+\s*[APM]{2})\s*</td>'
        $authorizationTime = ([regex]::Match($htmlBody, $authorizationTimePattern)).Groups[1].Value

        $authorizationDatePattern = 'Authorization Date\s*</td>\s*<td[^>]*>\s*([A-Za-z]{3}\s+\d{2},\s+\d{4})\s*</td>'
        $authorizationDate = ([regex]::Match($htmlBody, $authorizationDatePattern)).Groups[1].Value

        $referenceNumberPattern = 'Reference Number\s*</td>\s*<td[^>]*>\s*(\d+)\s*</td>'
        $referenceNumber = ([regex]::Match($htmlBody, $referenceNumberPattern)).Groups[1].Value

        $authorizationCodePattern = 'Authorization Code\s*</td>\s*<td[^>]*>\s*([A-Za-z0-9]+)\s*</td>'
        $authorizationCode = ([regex]::Match($htmlBody, $authorizationCodePattern)).Groups[1].Value

        $amountPattern = 'Amount\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $amount = ([regex]::Match($htmlBody, $amountPattern)).Groups[1].Value

        $cardEndingPattern = 'Card ending in.*?(\d{4})'
        $cardEnding = ([regex]::Match($htmlBody, $cardEndingPattern)).Groups[1].Value

        $calculatedPattern = 'Calculated\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $calculated = ([regex]::Match($htmlBody, $calculatedPattern)).Groups[1].Value

        $salesTaxPattern = 'Sales Tax\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $salesTax = ([regex]::Match($htmlBody, $salesTaxPattern)).Groups[1].Value

        $taxesFeesPattern = 'Taxes and Fees\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $taxesFees = ([regex]::Match($htmlBody, $taxesFeesPattern)).Groups[1].Value

        $subtotalPattern = 'Subtotal\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $subtotal = ([regex]::Match($htmlBody, $subtotalPattern)).Groups[1].Value

        $totalSavingsPattern = 'Total Savings\s*</td>\s*<td[^>]*>\s*-\$([\d.]+)\s*</td>'
        $totalSavings = ([regex]::Match($htmlBody, $totalSavingsPattern)).Groups[1].Value

        $totalItemsPattern = 'Total Items.*?(\d+)'
        $totalItems = ([regex]::Match($htmlBody, $totalItemsPattern)).Groups[1].Value

        $productPattern = '<tr>\s*<td[^>]*>\s*(.*?)\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>\s*</tr>\s*<tr>\s*<td[^>]*>\s*Quantity:\s*(\d+)\s*</td>'
        $products = [regex]::Matches($htmlBody, $productPattern)

        # Check for duplicates
        $checkQuery = "SELECT COUNT(*) FROM Receipts WHERE TransactionNumber = '$transactionNumber' AND ReferenceNumber = '$referenceNumber'"
        $recordset = $connection.Execute($checkQuery)
        $count = $recordset.Fields.Item(0).Value

        if ($count -eq 0) {
            foreach ($product in $products) {
                $productName = $product.Groups[1].Value
                $productPrice = $product.Groups[2].Value
                $productQuantity = $product.Groups[3].Value

                # Escape single quotes in the product name to prevent SQL errors
                $escapedProductName = $productName -replace "'", "''"
                $escapedProductPrice = $productPrice -replace "'", "''"
                $escapedProductQuantity = $productQuantity -replace "'", "''"

                $insertQuery = @"
INSERT INTO Receipts (
    Email, TransactionNumber, AuthorizationTime, AuthorizationDate, ReferenceNumber, AuthorizationCode,
    Amount, CardEndingIn, Calculated, SalesTax, TaxesAndFees, Subtotal, TotalSavings, TotalItems,
    ProductName, ProductPrice, ProductQuantity, ProductRegularPrice
) VALUES (
    '$($message.SenderEmailAddress)', '$transactionNumber', '$authorizationTime', 
    '$authorizationDate', '$referenceNumber', '$authorizationCode', 
    '$amount', '$cardEnding', '$calculated', '$salesTax', 
    '$taxesFees', '$subtotal', '$totalSavings', '$totalItems', 
    '$escapedProductName', '$escapedProductPrice', '$escapedProductQuantity', '$escapedProductPrice'
);
"@
                $connection.Execute($insertQuery) | Out-Null
            }
        } else {
            Write-Host "Duplicate record found for TransactionNumber: $transactionNumber, ReferenceNumber: $referenceNumber. Skipping insertion."
        }
    }
    $connection.Close()
    Write-Host "Data inserted successfully."
} else {
    Write-Host "No emails found in the specified folder."
}
