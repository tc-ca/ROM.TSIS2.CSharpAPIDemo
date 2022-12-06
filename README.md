# ROM.TSIS2.CSharpAPIDemo
A console app written in C# that shows how to communicate with the ROM TSIS2 API

# Developer Instructions
1. Open the project in Visual Studio and add a file to the project named **Secrets.config**
2. Enter the app settings by using this format:

```
<appSettings>
	<add key="Url" value="https://---.crm3.dynamics.com/" />
	<add key="ClientId" value="Enter_Your_Client_ID" />
	<add key="ClientSecret" value="Enter_Your_Secret" />
	<add key="TenantId" value="TenantID_From_Azure" />
	<add key="Authority" value="https://login.microsoftonline.com/---/oauth2/token" />
</appSettings>
```
3. Compile and run the app
