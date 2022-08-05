<##Author: Albano Lala
##Details: Graph / PowerShell Script cancellazione email prima di una certa data, 
##         Testare lo script prima della messa in produzione.
        .SYNOPSIS
        Cancellazione di tutte le email della mailbox interessata

        .PARAMETER Mailbox
        Utilizzare l'UPN della mailbox dedicata

        .PARAMETER ClientID
        Application (Client) ID dell'applicazione registrata su Azure AD

        .PARAMETER ClientSecret
        Il client secret creato dopo la registrazione dell'applicazione su Azure AD

        .PARAMETER TenantID
        L'ID del vostro Tenant

        .EXAMPLE
        .\graph-DeleteALLMails.ps1 -Mailbox "utente@contoso.com" -ClientSecret $clientSecret -ClientID $clientID -TenantID $tenantID
        
    
    #>

##

Param(
    [parameter(Mandatory = $true)]
    [String]
    $ClientSecret,
    [parameter(Mandatory = $true)]
    [String]
    $ClientID,
    [parameter(Mandatory = $true)]
    [String]
    $TenantID,
    [parameter(Mandatory = $true)]
    [String]
    $Mailbox
    )

##FUNZIONI##
function GetGraphToken {
    # Funzione dedicata alla richiesta del token di autenticazione e autorizzazione ad azure AD da parte dell'applicazione
    # Get OAuth token for a AAD Application (ci verrà dato in pasto un token che metteremo nella variabile $token)
    <#
        .SYNOPSIS
        Questa funzione fa richiesta e ottiene un token di autenticazione e autorizzazione da AAD utilizzando i seguenti paramentri (che dovrete fornire voi in fase di esecuzione dello script)
    
        .PARAMETER clientSecret
        - è il client secret dell'applicazione di che avete creato in Azure AD (il client secret dovrete crearlo all'interno dell'applicazione dopo aver creato quest'ultima)
    
        .PARAMETER clientID
        - è il client ID dell'applicazione che avete creato su AAD
        .PARAMETER tenantID
        - è il Tenant ID del vostro TENANT
        
        #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $ClientSecret,
        [parameter(Mandatory = $true)]
        [String]
        $ClientID,
        [parameter(Mandatory = $true)]
        [String]
        $TenantID
    
    )
    
        
        
    # Costruzione URI 
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
         
    # Costruzione del body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
         
    # Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
         
    # Memorizziamo il contenuto dell'Access Token ricevuto all'iterno della variabile $token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
    return $token
}

function RunQueryandEnumerateResults {
    <#
    .SYNOPSIS
    La function esegue query su Graph e, se sono presenti pagine aggiuntive, le analizza e le aggiunge ad una singola variabile
    
    .PARAMETER apiUri
    -APIURi è l'apiUri da passare
    
    .PARAMETER token
    -token è il token di autenticazione
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $apiUri,
        [parameter(Mandatory = $true)]
        $token

    )

    #Query Graph
    $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
    #write-host $results

    #Inizio popolazione risultati
    $ResultsValue = $Results.value

    #Se è presente una pagina successiva, interroga la pagina successiva finché non ci sono più pagine e aggiungi i risultati al set esistente
    if ($null -ne $results."@odata.nextLink") {
        write-host enumerating pages -ForegroundColor yellow
        $NextPageUri = $results."@odata.nextLink"
        ##Finche c'è una pagina successiva, esegui una query e un ciclo e aggiungi i risultati
        While ($null -ne $NextPageUri) {
            $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $NextPageURI -Method Get)
            $NxtPageData = $NextPageRequest.Value
            $NextPageUri = $NextPageRequest."@odata.nextLink"
            #mettiamo in tail i risultati delle pagine
            $ResultsValue = $ResultsValue + $NxtPageData
        }
    }

    ##La function ritorna i risultati delle query fatte tramite delegation di graph api
    return $ResultsValue

    
}

function DeleteMail {
    <#
.SYNOPSIS
Cancellazione email dalla mailbox

.PARAMETER mail
ID della mail da cancellare

.PARAMETER token
Access token ottenuto precedentemente tramite graph, con questo token possiamo accedere in read/write alla mailbox interessata

.PARAMETER mailbox
UPN dell'utente

#>
    Param(
        [parameter(Mandatory = $true)]
        $mail,
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $mailbox

    )

    $Apiuri = "https://graph.microsoft.com/v1.0/users/$mailbox/messages/$mail/move"

    $Destination = @"
    {
        "destinationId": "recoverableitemspurges"
    }
"@

(Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -ContentType 'application/json' -Body $destination -Uri $apiUri -Method Post)

}


$token = GetGraphToken -ClientSecret $ClientSecret -ClientID $ClientID -TenantID $TenantID
$mailbox = $mailbox.replace("'","`'")
$Apiuri = "https://graph.microsoft.com/v1.0/users/$mailbox/messages"

write-host "checking Mails via: $Apiuri"
$results = RunQueryandEnumerateResults -apiUri $apiuri -token $token

write-host "Found $($results.count) mails"

foreach ($mail in $Results) {

    write-host "Processing $($mail.subject)"

    DeleteMail -token $token -mail $mail.id -mailbox $mailbox
}