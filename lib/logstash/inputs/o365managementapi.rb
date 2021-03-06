# encoding: utf-8
require "logstash/inputs/base"
require "logstash/namespace"
require 'logstash/inputs/o365managementapi/O365ManagementapiHelper'

# Connector for Office 365 anagement API logs.
#
# This plugin is intented to download last minutes logs periodically.
#
# ==== Usage:
#
# Here is an example of how to configure the plugin.
# First, you need to register your application in Azure AD and get access tokens.
# Then, generate a self signed certificate to enable service-to-service calls and 
# add the public key thumbprint in your Azure application manifest.
# Check https://docs.microsoft.com/en-us/office/office-365-management-api/get-started-with-office-365-management-apis
# Once you have done this get you client ID and Tenant ID in Azure.
#
# In this example, we retrieve the last 12 minutes o365 Sharepoint logs every 10 minutes. 
# It is very important to have overlapping query periods to make sure you are not missing logs.
# This also means that some logs will be sent twice to the output. Make sure to map Document id
# to log id in elasticsearch to prevent duplicates.
#
# [source,ruby]
# ------------------------------------------------------------------------------
# input {
#   O365managementapi {
#     clientid => "XXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
#     tenantid => "XXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
#     tenant => "XXXX.onmicrosoft.com"
#     pfx_file => "o365.pfx"
#     pfx_password => "secret"
#     timerange => 12
#     schedule => "*/10 * * * *"
#     content_type => "Audit.SharePoint"
#   }
# }
# ------------------------------------------------------------------------------
#
# Very important : 
#
# In order to prevent duplicates in elasticsearch, don't forget to specify 'Id' field as
# document id :
#
# output {
#  elasticsearch {
#    hosts => "example.com"
#    document_id => "%{[Id]}"
#  }
#}
#

class LogStash::Inputs::O365managementapi < LogStash::Inputs::Base
  config_name "o365managementapi"

  # If undefined, Logstash will complain, even if codec is unused.
  default :codec, "plain"

  # Azure application client id
  config :clientid, :validate => :string, :required => true
  
  # Your azure tenant id
  config :tenantid, :validate => :string, :required => true
  
  # Your azure tenant
  config :tenant, :validate => :string, :required => true

   # Publisher ID. default : tenantid
  config :publisherid, :validate => :string, :required => false
  
  # File path of the certificate used for service-to-service API call.
  config :pfx_file, :validate => :path, :required => true
  
  # Certificate password.
  config :pfx_password, :validate => :string, :required => true
  
  # Last n minutes of content to download.
  # for example: 25 (get logs of last 25 minutes)
  #
  # The default, `10`, means downloading last 10 minutes logs.
  config :timerange, :validate => :number, :default => 10
  
  # Schedule of when to periodically dowload o365 logs, in Cron format
  # for example: "*/20 * * * *" (execute query every 30 minutes)
  #
  # There is no schedule by default. If no schedule is given, then the query is run
  # exactly once.
  config :schedule, :validate => :string
 
  # Used for one-time 24h data port. If specified, timerange is ignored. If a schedule is specified, this parameter is ignored.
  # for example: "2018-11-07"
  config :import_date, :validate =>:string

  # Used for one-time 24h data port. If specified, timerange is ignored. If a schedule is specified, this parameter is ignored.
  # for example: 30
  config :days_to_import, :validate => :number, :default => 1
 
  # Content type. Can be "Audit.AzureActiveDirectory", "Audit.Exchange", "Audit.SharePoint", "Audit.General" or "DLP.All"
  # for more information on content types see https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference
  config :content_type, :validate => :string, :required => true

  public
  def register
	require "rufus/scheduler"
	
	# if not defined, publisher ID = tenant ID
	unless @publisherid
		@publisherid = @tenantid
	end
	@helper = O365ManagementapiHelper.new(@tenant, @tenantid, @publisherid, @clientid, @pfx_file, @pfx_password, @content_type, logger)
  end # def register

  def run(queue)
    sleep(2)
    #Manage subscription
    subscribed = @helper.subscribed()
    logger.info("[#{@content_type}] Already subscribed? #{subscribed}")
    unless subscribed
        logger.info("[#{@content_type}] subscribing...")
	@helper.subscribe
	logger.info("[#{@content_type}] subscribed.")
    end
    if @schedule
	logger.info("[#{@content_type}] Schedule defined, initiating scheduler with schedule #{@schedule}")
	@scheduler = Rufus::Scheduler.new(:max_work_threads => 1)
	@scheduler.cron @schedule do
		process(queue)
	end
	@scheduler.join
    else
	logger.info("[#{@content_type}] no scheduler defined, running once.")
      	process(queue)
    end
  end # def run

  private
  def process(queue)
	start_time = ""
	end_time = ""
	#one time import, by 24h querries
	if !@schedule and @import_date
		$i = @days_to_import - 1
		while $i >= 0
			end_ = DateTime.strptime(@import_date + "T23:59", "%Y-%m-%dT%H:%M") - $i
			start = DateTime.strptime(@import_date + "T23:59", "%Y-%m-%dT%H:%M") - $i - 1
			start_time = start.strftime("%Y-%m-%dT%H:%M")
			end_time = end_.strftime("%Y-%m-%dT%H:%M")
			logs = @helper.get_logs(start_time, end_time)
                	logs.each do |log|
                        	event = LogStash::Event.new(log)
                        	decorate(event)
                        	queue << event
                	end
			$i -= 1	
		end
	else
		now = Time.now.utc
		start = now - (@timerange*60)
		start_time = start.strftime("%Y-%m-%dT%H:%M")
		end_time = now.strftime("%Y-%m-%dT%H:%M")
		logs = @helper.get_logs(start_time, end_time)
		logs.each do |log|
			event = LogStash::Event.new(log)
			decorate(event)
			queue << event
		end
	end
  end # process(queue)
  
  def stop
	@scheduler.shutdown(:wait) if @scheduler
  end
end # class LogStash::Inputs::O365managementapi
