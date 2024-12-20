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
# I will output a CSV to your desktop by default to a file call extracted_data.csv - You can change the path in the variable just below.
#
#set the output
$outputPath = "$([Environment]::GetFolderPath('Desktop'))\extracted_data.csv"

# Create the Outlook application object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Get the Inbox folder for your Gmail account (it should be listed as one of your email accounts in Outlook)
$inbox = $namespace.Folders.Item("account name")  # Replace with your account name
$gmailFolder = $inbox.Folders.Item("Inbox")  # Access the Inbox folder of Gmail
$messages = $gmailFolder.Items  # Get all messages in the Inbox

# Ensure there are at least three emails
if ($messages.Count -ge 1) {
    $data = @()
    for ($i = 1; $i -le $messages.Count; $i++) {
        # Get the message from the inbox
        $message = $messages.Item($i)

        # Extract basic details from the message
        $subject = $message.Subject
        $receivedTime = $message.ReceivedTime
        $htmlBody = $message.HTMLBody
        
        Write-Host "Processing Email $i`:"
        Write-Host "Subject: $subject"
        Write-Host "Received: $receivedTime"
        
        # Regex patterns for the fields we need to extract:
        
        # Transaction Number (adjust based on email format)
        $transactionNumberPattern = 'Transaction Number\s*</td>\s*<td[^>]*>\s*([\d\s]+)\s*</td>'
        $transactionNumberMatch = [regex]::Match($htmlBody, $transactionNumberPattern)
        $transactionNumber = if ($transactionNumberMatch.Success) { $transactionNumberMatch.Groups[1].Value.Replace(' ', '') } else { "Not found" }

        # Authorization Time (adjust based on email format)
        $authorizationTimePattern = 'Authorization Time\s*</td>\s*<td[^>]*>\s*([\d:]+\s*[APM]{2})\s*</td>'
        $authorizationTimeMatch = [regex]::Match($htmlBody, $authorizationTimePattern)
        $authorizationTime = if ($authorizationTimeMatch.Success) { $authorizationTimeMatch.Groups[1].Value } else { "Not found" }
        
        # Authorization Date (adjust based on email format)
        $authorizationDatePattern = 'Authorization Date\s*</td>\s*<td[^>]*>\s*([A-Za-z]{3}\s+\d{2},\s+\d{4})\s*</td>'
        $authorizationDateMatch = [regex]::Match($htmlBody, $authorizationDatePattern)
        $authorizationDate = if ($authorizationDateMatch.Success) { $authorizationDateMatch.Groups[1].Value } else { "Not found" }
        
        # Reference Number (adjust based on email format)
        $referenceNumberPattern = 'Reference Number\s*</td>\s*<td[^>]*>\s*(\d+)\s*</td>'
        $referenceNumberMatch = [regex]::Match($htmlBody, $referenceNumberPattern)
        $referenceNumber = if ($referenceNumberMatch.Success) { $referenceNumberMatch.Groups[1].Value } else { "Not found" }
        
        # Authorization Code (adjust based on email format)
        $authorizationCodePattern = 'Authorization Code\s*</td>\s*<td[^>]*>\s*([A-Za-z0-9]+)\s*</td>'
        $authorizationCodeMatch = [regex]::Match($htmlBody, $authorizationCodePattern)
        $authorizationCode = if ($authorizationCodeMatch.Success) { $authorizationCodeMatch.Groups[1].Value } else { "Not found" }
        
        # Amount (adjust based on email format)
        $amountPattern = 'Amount\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $amountMatch = [regex]::Match($htmlBody, $amountPattern)
        $amount = if ($amountMatch.Success) { $amountMatch.Groups[1].Value } else { "Not found" }
        
        # Card ending in (adjust based on email format)
        $cardEndingPattern = 'Card ending in.*?(\d{4})'
        $cardEndingMatch = [regex]::Match($htmlBody, $cardEndingPattern)
        $cardEnding = if ($cardEndingMatch.Success) { $cardEndingMatch.Groups[1].Value } else { "Not found" }
        
        # Calculated (adjust based on email format)
        $calculatedPattern = 'Calculated\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $calculatedMatch = [regex]::Match($htmlBody, $calculatedPattern)
        $calculated = if ($calculatedMatch.Success) { $calculatedMatch.Groups[1].Value } else { "Not found" }
        
        # Sales Tax (adjust based on email format)
        $salesTaxPattern = 'Sales Tax\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $salesTaxMatch = [regex]::Match($htmlBody, $salesTaxPattern)
        $salesTax = if ($salesTaxMatch.Success) { $salesTaxMatch.Groups[1].Value } else { "Not found" }
        
        # Taxes and Fees (adjust based on email format)
        $taxesFeesPattern = 'Taxes and Fees\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $taxesFeesMatch = [regex]::Match($htmlBody, $taxesFeesPattern)
        $taxesFees = if ($taxesFeesMatch.Success) { $taxesFeesMatch.Groups[1].Value } else { "Not found" }
        
        # Subtotal (adjust based on email format)
        $subtotalPattern = 'Subtotal\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>'
        $subtotalMatch = [regex]::Match($htmlBody, $subtotalPattern)
        $subtotal = if ($subtotalMatch.Success) { $subtotalMatch.Groups[1].Value } else { "Not found" }
        
        # Total Savings (adjust based on email format)
        $totalSavingsPattern = 'Total Savings\s*</td>\s*<td[^>]*>\s*-\$([\d.]+)\s*</td>'
        $totalSavingsMatch = [regex]::Match($htmlBody, $totalSavingsPattern)
        $totalSavings = if ($totalSavingsMatch.Success) { $totalSavingsMatch.Groups[1].Value } else { "Not found" }
        
        # Total Items (adjust based on email format)
        $totalItemsPattern = 'Total Items.*?(\d+)'
        $totalItemsMatch = [regex]::Match($htmlBody, $totalItemsPattern)
        $totalItems = if ($totalItemsMatch.Success) { $totalItemsMatch.Groups[1].Value } else { "Not found" }

        # Extract product details (adjust based on email format)
        $productPattern = '<tr>\s*<td[^>]*>\s*(.*?)\s*</td>\s*<td[^>]*>\s*\$([\d.]+)\s*</td>\s*</tr>\s*<tr>\s*<td[^>]*>\s*Quantity:\s*(\d+)\s*</td>\s*</tr>(?:\s*<tr>\s*<td[^>]*>\s*Regular Price\s*\$\s*([\d.]+)\s*</td>\s*</tr>)?'
        $matches = [regex]::Matches($htmlBody, $productPattern)
        $productList = @()
        foreach ($match in $matches) {
            $productName = $match.Groups[1].Value.Trim()
            $productPrice = $match.Groups[2].Value.Trim()
            $productQuantity = $match.Groups[3].Value.Trim()
            $productRegularPrice = if ($match.Groups[4].Success) { $match.Groups[4].Value.Trim() } else { "N/A" }
            $productList += [PSCustomObject]@{
                Name = $productName
                Price = $productPrice
                Quantity = $productQuantity
                RegularPrice = $productRegularPrice
            }
        }

        # Add extracted details to data array
        $data += [PSCustomObject]@{
            Email = $i
            TransactionNumber = $transactionNumber
            AuthorizationTime = $authorizationTime
            AuthorizationDate = $authorizationDate
            ReferenceNumber = $referenceNumber
            AuthorizationCode = $authorizationCode
            Amount = $amount
            CardEndingIn = $cardEnding
            Calculated = $calculated
            SalesTax = $salesTax
            TaxesAndFees = $taxesFees
            Subtotal = $subtotal
            TotalSavings = $totalSavings
            TotalItems = $totalItems
            Products = $productList
        }
    }}
# Save data to CSV
$headers = @("Email", "Transaction Number", "Authorization Time", "Authorization Date", "Reference Number", "Authorization Code", "Amount", "Card Ending In", "Calculated", "Sales Tax", "Taxes and Fees", "Subtotal", "Total Savings", "Total Items", "Product Name", "Product Price", "Product Quantity", "Product Regular Price")

# Initialize an array to hold the flattened data
$flattenedData = @()

foreach ($emailData in $data) {
    foreach ($product in $emailData.Products) {
        $flattenedData += [PSCustomObject]@{
            Email = $emailData.Email
            TransactionNumber = $emailData.TransactionNumber
            AuthorizationTime = $emailData.AuthorizationTime
            AuthorizationDate = $emailData.AuthorizationDate
            ReferenceNumber = $emailData.ReferenceNumber
            AuthorizationCode = $emailData.AuthorizationCode
            Amount = $emailData.Amount
            CardEndingIn = $emailData.CardEndingIn
            Calculated = $emailData.Calculated
            SalesTax = $emailData.SalesTax
            TaxesAndFees = $emailData.TaxesAndFees
            Subtotal = $emailData.Subtotal
            TotalSavings = $emailData.TotalSavings
            TotalItems = $emailData.TotalItems
            ProductName = $product.Name
            ProductPrice = $product.Price
            ProductQuantity = $product.Quantity
            ProductRegularPrice = $product.RegularPrice
        }
    }
}

$flattenedData | Export-Csv -Path $outputPath -NoTypeInformation

Write-Host "Data has been successfully written to $outputPath."
