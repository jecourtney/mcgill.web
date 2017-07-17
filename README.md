-------------------------
Local development machine
-------------------------
c:\d-drive\project\core\McGill.Web\code .
c:\d-drive\project\core\McGillWebAPI\code .
c:\d-drive\project\angular2\unitedmcgill\code .

-----------------
Start server code
-----------------
c:\d-drive\project\core\McGillWebAPI\dotnet run
debug from inside vscode w/out dotnet run

-----------------
Client code
-----------------
c:\d-drive\project\angular2\unitedmcgill\ng serve
debug from inside vscode w/ ng serve

----------------
deploy
----------------
dotnet publish -f netcoreapp1.1 -c Release
Delete destination folder contents
Copy contents of bin\release\netcoreapp1.1\publish to destination IIS folder (see below how to setup)
// if you want to run in a command prompt, dotnet mcgillwebapi.dll from the server folder


---------------------------------
Host ASP.NET CORE with IIS
---------------------------------
Install the .NET Core Windows Server Hosting bundle:  https://go.microsoft.com/fwlink/?linkid=844461
This will give a AspNetCoreModule in the Modules of IIS
publish the ASP.NET core:  dotnet publish -f netcoreapp1.1 -c Release and copy the output to the server
create a new IIS Application Pool for AspDotNetCore if you don't already have one, set to No Managed Code
Create your new application under an existing site or create a new one and add the application to it
Set the Physical path to the output folder on the server you created
set bindings and start the site it will create a reverse proxy to your application


----------------------------------
Use weblistener instead of kestrel
----------------------------------
** STUCK HERE CORS is not working in WebListener **
Add  
    <PackageReference Include="Microsoft.AspNetCore.Server.WebListener" Version="1.1.2" />
    <PackageReference Include="Microsoft.Net.Http.Server" Version="1.1.2" />
to .csproj

Call the UseWebListener extension method on WebHostBuilder in your Main method, specifying any WebListener options and settings that you need in 

Program.cs
                .UseWebListener(options =>
                {
                        options.ListenerSettings.Authentication.Schemes = AuthenticationSchemes.None;
                        options.ListenerSettings.Authentication.AllowAnonymous = true;
                })
Preregister URL prefixes for the port you want to use
netsh http show urlacl
Make sure the one you want to reserve is available, if not, delete the existing one
netsh http delete urlacl url=https://+:443/C574AC30-5794-4AEE-B1BB-6651C5315029/
netsh http delete urlacl url=https://+:443/sra_{BA195980-CD49-458b-9E23-C84EE0ADCD75}/
netsh http add urlacl url=http://+:443/ user=jcourtney
** STUCK HERE CORS is not working in WebListener **


---------
etc
---------

dotnet restore
dotnet run
dotnet watch
dotnet pack (makes a nuget package) that can be copied into c:\nuget and used by your other projects by adding them to project.json

$env:ASPNETCORE_ENVIRONMENT = "Development"
dotnet ef migrations add ...
dotnet ef database

-------------------------------
Yoman - generator omnisharp.net
-------------------------------
npm install -g generator
yo aspnet
	webapi

cd Controller
yo aspnet:webapicontroller

cd Models
yo aspnet:Class

----------------------------
nuget
----------------------------
nuget sources add -name LocalPackages -source "c:\nuget"


-----------------------------
github
-----------------------------
echo "# mcgill.web" >> README.md
git init
git add README.md
git commit -m "first commit"
git remote add origin https://github.com/jecourtney/mcgill.web.git
git push -u origin master
