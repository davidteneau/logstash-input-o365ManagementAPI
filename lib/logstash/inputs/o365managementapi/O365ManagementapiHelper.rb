require 'adal'
require 'openssl'
require 'requests'
require 'json'
require 'parallel'

class O365ManagementapiHelper
	ADAL::TokenRequest::GrantType::JWT_BEARER = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
	def initialize(tenant, tenantid, publisherid, clientid, pfx_path, pfx_password, content_type, logger)
		@tenant = tenant
		@tenantid = tenantid
		@publisherid = publisherid
		@clientid = clientid
		@pfx_path = pfx_path
		@pfx_password = pfx_password
		@content_type = content_type
		@adal_response = nil
		@authentication_context = nil
		@client_cred = nil
		@logger = logger
		
		pfx = OpenSSL::PKCS12.new(File.read(@pfx_path), @pfx_password)
		authority = ADAL::Authority.new(ADAL::Authority::WORLD_WIDE_AUTHORITY, @tenant)
		@client_cred = ADAL::ClientAssertionCertificate.new(authority, @clientid, pfx)
		@authentication_context = ADAL::AuthenticationContext
                        .new(ADAL::Authority::WORLD_WIDE_AUTHORITY, @tenant)
		@adal_response = @authentication_context.acquire_token_for_client('https://manage.office.com', @client_cred)
		case @adal_response
			when ADAL::SuccessResponse
				@logger.info("[#{@content_type}] Successfully got Access Token. Expires in #{@adal_response.expires_in} seconds")
			when ADAL::FailureResponse
				@logger.error("[#{@content_type}] Failed to authenticate with client credentials. Received error: " \
				"#{@adal_response.error}\n and error description: #{@adal_response.error_description}.")
		end
	end #initialize

	def refresh_token_if_needed()
		expires_in = @adal_response.expires_on.to_i - Time.now.to_i
		@logger.info("[#{@content_type}] Token expires in #{expires_in} seconds")
		if expires_in <= 0
			 @logger.info("[#{@content_type}] Token expired, or about to expire. Acquiring new token.")
			 response = @authentication_context.acquire_token_for_client('https://manage.office.com', @client_cred)
			 case response
                                when ADAL::SuccessResponse
                                        @logger.info("[#{@content_type}] Successfully got Access Token. Expires in #{response.expires_in} seconds)")
                                        @adal_response = response
                                when ADAL::FailureResponse
                                        @logger.error("[#{@content_type}] Failed to authenticate with client credentials. Received error: " \
                                        "#{response.error} and error description: #{response.error_description}.")
                        end
		else
                        @logger.info("[#{@content_type}] Token validity is ok, no need to refresh it.")
		end
		
	end #refresh_token_if_needed()
	
	# Check if the subscription is active for this content type
	def subscribed()
		#add access token to header
		header = {"Authorization": "Bearer #{@adal_response.access_token}"}
		#poll to get available content
		request_uri = "https://manage.office.com/api/v1.0/#{@tenantid}/activity/feed/subscriptions/list"
                response = Requests.request("GET", request_uri, headers: header)
                @logger.info("[#{@content_type}] Subscription list: \n#{response.json()}}")

		response.json().each do |item|
			if item["contentType"] == @content_type
				return true
			end
		end
		false
	end
	
	# get logs from Office365 Management API. start time and end time in format "%Y-%m-%dT%H:%M")
	def get_logs(start_time, end_time)
		refresh_token_if_needed()
		logs = Array.new
		#add access token to header
		header = {"Authorization": "Bearer #{@adal_response.access_token}", "Content-Length": "0"}
    		request_uri = "https://manage.office.com/api/v1.0/#{@tenantid}/activity/feed/subscriptions/content?contentType=#{@content_type}&startTime=#{start_time}&endTime=#{end_time}&PublisherIdentifier=#{@publisherid}"
    		# Send request to get blobs uri
    		begin
                        response = Requests.request("GET",request_uri, headers: header)
			while response.headers.include? 'nextpageuri'
                        	next_page = response.headers['nextpageuri'][0]
                        	process_page(response, logs)
                        	#use next page uri to get more blobs
                        	response = Requests.request("GET", "#{next_page}?PublisherIdentifier=#{@publisherid}", headers: header)
                	end
                rescue StandardError => e
                        @logger.error("[#{@content_type}] Error getting page #{request_uri}\n#{e.message}")
                        @logger.error(e.backtrace.inspect)
                end
    		process_page(response, logs)
		@logger.info("[#{@content_type}] Processed #{logs.count} logs, start:#{start_time}, end:#{end_time}")
    		logs
	end
	
	# Process logs for one page. Private method
	def process_page(response, logs)
		blobs = response.json()
    		uri_list = Array.new
    		blobs.each do |blob|
      			uri_list.push(blob['contentUri'])
    		end
		Parallel.each(uri_list, in_threads: uri_list.count) do |uri|
    		#Parallel.each(uri_list, in_threads: 10) do |uri|
      			process_blob(uri, logs)
    		end
	end
	
   	# Process logs of one blob. Private Method
	def process_blob(uri, logs)
		request_uri = "#{uri}?PublisherIdentifier=#{@publisherid}"
		header = {"Authorization": "Bearer #{@adal_response.access_token}", "Content-Length": "0"}
	 	#puts "header: #{header}" 
    		#poll to get available content
		begin
    			event_blob = Requests.request("GET",request_uri, headers: header)
			if event_blob.status == 200
      				events = JSON.parse(event_blob.body)
      				#Iterate over bblob content and push each event to logs
      				events.each do |event|
        				logs << event
      				end
			end
		rescue StandardError => e
			@logger.error("[#{@content_type}] Error retreiving blob content: #{e.message}")
			@logger.error("[#{@content_type}] " + e.backtrace.inspect)
			@logger.error("[#{@content_type}]  URI: #{request_uri}")
		end
	end
	
	# start a subscription
  	def subscribe()
     		#add access token to header
    		header = {"Authorization": "Bearer #{@adal_response.access_token}", "Content-Length": "0"}
    		request_uri = "https://manage.office.com/api/v1.0/#{@tenantid}/activity/feed/subscriptions/start?contentType=#{@content_type}&PublisherIdentifier=#{@publisherid}"
    		# start subscription
    		begin
                        response = Requests.request("POST",request_uri, headers: header)
                rescue StandardError => e
                        @logger.error("[#{@content_type}] Error subscribing: #{e.message}")
                        @logger.error("[#{@content_type}] " + e.backtrace.inspect)
     		end
		if response.status == 200
      			@logger.info("[#{@content_type}] sucessfully subscribed:\n#{response.body}")
			return true
    		end
    		#default : return false
    		false
  	end
  
  	# stop a subscription
  	def unsubscribe()
     		#add access token to header
    		header = {"Authorization": "Bearer #{@adal_response.access_token}", "Content-Length": "0"}
    		request_uri = "https://manage.office.com/api/v1.0/#{@tenantid}/activity/feed/subscriptions/stop?contentType=#{@content_type}&PublisherIdentifier=#{@publisherid}"
    		# start subscription
    		begin
                        response = Requests.request("POST",request_uri, headers: header)
                rescue StandardError => e
                        @logger.error("[#{@content_type}] Error unsubscribing: #{e.message}")
                        @logger.error("[#{@content_type}] " + e.backtrace.inspect)
                end
		if response.status == 200
      			return true
    		end
    		#default : return false
    		false
  	end
  
  	private :process_page; :process_blob
  
end # class o365managementapiHelper

