# SA-UnityEdms
Web Service API wrapping access to OnBase Unity AppServer


Required config file values:

<appSettings>
  <add key="OnBaseUrl" value="http://sampleuri.com/AppServer/service.asmx"/>
  <add key="OnBaseDataSource" value="myDataSource"/>
</appSettings>

<connectionStrings>
  <add name="OnBase" connectionString="server=myDataSource;uid=myUser;password=myPass;" />
</connectionStrings>


The wrapper only supports Oracle DB backend.
