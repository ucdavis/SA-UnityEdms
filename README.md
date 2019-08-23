# SA-UnityEdms
Web Service API wrapping access to OnBase Unity AppServer


Required config file values:

<pre>
&lt;appSettings&gt;
  &lt;add key="OnBaseUrl" value="http://sampleuri.com/AppServer/service.asmx"/&gt;
  &lt;add key="OnBaseDataSource" value="myDataSource"/&gt;
&lt;/appSettings&gt;

&lt;connectionStrings&gt;
  &lt;add name="OnBase" connectionString="server=myDataSource;uid=myUser;password=myPass;" /&gt;
&lt;/connectionStrings&gt;
</pre>

The wrapper only supports SQL Server backend.Make sure that the database username and password you specify have write access to the logging tables.
