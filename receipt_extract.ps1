
#set the email address
$emailaccountname = "account.name@gmail.com"


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
