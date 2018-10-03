# Logstash input Plugin for Office 365 Audit logs

This is an input plugin for [Logstash](https://github.com/elastic/logstash).

It is intented to download last minutes office 365 logs periodically.

#### Install
- To get started, you'll need JRuby with the Bundler gem installed.

- clone this repo :
```sh
git clone https://github.com/davidteneau/logstash-input-o365managementapi.git
cd logstash-input-o365managementapi
```

- Install dependencies
```sh
bundle install
```

- Build the plugin gem
```sh
gem build logstash-input-o365managementapi.gemspec
```

- Install the plugin from the Logstash home
```sh
bin/logstash-plugin install /path/to/logstash-input-o365managementapi/logstash-input-o365managementapi-1.0.1.gem
```
#### Usage:
Here is an example of how to configure the plugin.

First, you need to register your application in Azure AD and get access tokens.
Then, generate a self signed certificate to enable service-to-service calls and
add the public key thumbprint in your Azure application manifest.
Check [Microsoft documentation](https://docs.microsoft.com/en-us/office/office-365-management-api/get-started-with-office-365-management-apis)
Once you have done this retrieve from Azure you client ID and Tenant information.

In this example, we retrieve the last 12 minutes o365 Sharepoint logs every 10 minutes.
It is very important to have overlapping query periods to make sure you are not missing logs.
This also means that some logs will be sent twice to the output. Make sure to map Document id
to log id in elasticsearch to prevent duplicates.

```Ruby
input {
   O365managementapi {
     clientid => "XXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
     tenantid => "XXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
     tenant => "XXXX.onmicrosoft.com"
     pfx_file => "o365.pfx"
     pfx_password => "secret"
     timerange => 12
     schedule => "*/10 * * * *"
     content_type => "Audit.SharePoint"
   }
}
```

Very important :

 In order to prevent duplicates in elasticsearch, don't forget to specify 'Id' field as document id :
```Ruby
 output {
  elasticsearch {
    hosts => "example.com"
    document_id => "%{[Id]}"
  }
}
```
