#dotnet tool install --global dotnet-ef; `
git clone https://github.com/Azure/Commercial-Marketplace-SaaS-Accelerator.git -b 7.1.0 --depth 1; `
cd ./Commercial-Marketplace-SaaS-Accelerator/deployment; `
.\Deploy.ps1 `
 -WebAppNamePrefix "ds07" `
 -ResourceGroupForDeployment "ds07-saas-accelerator" `
 -PublisherAdminUsers "admin@aptero.co" `
 -Location "West Europe" 
 
 033e8d47-3604-4d1d-ae30-d9567481b43f
 
 ✅ If the intallation completed without error complete the folllowing checklist:
   🔵 Add The following URL in PartnerCenter SaaS Technical Configuration
      ➡️ Landing Page section:       https://ds07-portal.azurewebsites.net/
      ➡️ Connection Webhook section: https://ds07-portal.azurewebsites.net/api/AzureWebhook
      ➡️ Tenant ID:                  48caed5d-d20d-40ce-b0cd-9261fc5d8f83
      ➡️ AAD Application ID section: b19f2f2b-83e2-4c04-a2db-810c0bb8ebcb
      
      
https://ds07-admin.azurewebsites.net/


Configure your application
Configure your application with the settings below.

Server	smtp.sendgrid.net
Ports	
25, 587	(for unencrypted/TLS connections)
465	(for SSL connections)
Username	apikey
Password	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