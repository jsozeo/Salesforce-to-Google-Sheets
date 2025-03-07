// Main function to execute the script
function main() {
    try {
      console.log('=== Starting script execution ===');
      
      // First test the API connection
      console.log('\nTesting API connection...');
      const testResult = testApiConnection();
      if (!testResult.success) {
        throw new Error('API connection test failed');
      }
      console.log('API connection test successful');
      
      // Then get the report data
      console.log('\nFetching report data...');
      const reportData = getSalesforceReport();
      
      console.log('=== Script execution completed successfully ===');
      return reportData;
    } catch (error) {
      console.error('=== Script execution failed ===');
      console.error('Error:', error.toString());
      throw error;
    }
} 


// Function to get OAuth 2.0 access token using Client Credentials grant
function getAccessToken() {
    console.log('Starting OAuth token acquisition...');
    
    const payload = {
      grant_type: 'client_credentials'
    };
    
    let options = {
      method: 'post',
      muteHttpExceptions: true
    };
    
    // Add client authentication based on configuration
    if (CONFIG.sendCredentialsInBody === 'true') {
      console.log('Using body authentication method');
      payload.client_id = CONFIG.clientId;
      payload.client_secret = CONFIG.clientSecret;
      options.payload = payload;
      console.log('Client credentials added to request body');
    } else {
      console.log('Using header authentication method');
      // Send credentials in Authorization header
      const credentials = Utilities.base64Encode(CONFIG.clientId + ':' + CONFIG.clientSecret);
      options.headers = {
        'Authorization': 'Basic ' + credentials
      };
      options.payload = payload;
      console.log('Client credentials added to Authorization header');
    }
    
    try {
      console.log('Sending token request to:', CONFIG.accessTokenUrl);
      const response = UrlFetchApp.fetch(CONFIG.accessTokenUrl, options);
      console.log('Token request response received');
      
      const responseText = response.getContentText();
      console.log('Token response text');
      
      const result = JSON.parse(responseText);
      console.log('\n=== TOKEN INFORMATION ===');
      console.log('Access Token:', result.access_token);
      console.log('Token Type:', result.token_type);
      console.log('Instance URL:', result.instance_url);
      console.log('Scope:', result.scope);
      console.log('ID:', result.id);
      console.log('Issued At:', new Date(parseInt(result.issued_at)).toISOString());
      console.log('=== END TOKEN INFORMATION ===\n');
      
      if (result.access_token) {
        console.log('Access token obtained successfully');
        return result;
      } else {
        console.error('No access token found in response');
        throw new Error('No access token in response');
      }
    } catch (error) {
      console.error('Error getting access token:', error.toString());
      throw error;
    }
}

// Function to fetch Salesforce report data
function getSalesforceReport() {
    try {
      console.log('Starting Salesforce report retrieval...');
      const tokenInfo = getAccessToken();
      console.log('Successfully obtained access token:', tokenInfo.access_token.substring(0, 10) + '...');
      
      // Use instance_url from token response, falling back to config if not available
      const baseUrl = tokenInfo.instance_url || CONFIG.salesforceBaseUrl;
      const salesforceUrl = baseUrl + CONFIG.analyticsReportUrl;
      console.log('Preparing request to:', salesforceUrl);
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${tokenInfo.access_token}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      console.log('Request headers:', JSON.stringify(options.headers));
      console.log('Sending request to Salesforce API...');
      
      const response = UrlFetchApp.fetch(salesforceUrl, options);
      const responseCode = response.getResponseCode();
      console.log('Response status:', responseCode);
      
      if (responseCode !== 200) {
        const errorText = response.getContentText();
        console.error('Error response:', errorText);
        throw new Error(`API request failed with status ${responseCode}: ${errorText}`);
      }
      
      const responseText = response.getContentText();
      console.log('Raw response length:', responseText.length);
      console.log('Raw response:', responseText);
      
      const result = JSON.parse(responseText);
      console.log('Response successfully parsed as JSON');
      
      // Log the full response body in a formatted way
      console.log('\nFormatted Response Body:');
      console.log(JSON.stringify(result, null, 2));
      
      // Also log specific sections of interest
      console.log('\nReport Details:');
      console.log('Report Name:', result.attributes?.reportName || 'Not found');
      console.log('Total Records:', result.factMap?.['T!T']?.aggregates?.[0]?.value || 'Not found');
      
      return result;
    } catch (error) {
      console.error('Error fetching Salesforce report:', error.toString());
      if (error.stack) {
        console.error('Stack trace:', error.stack);
      }
      throw error;
    }
}

// Function to test API connection
function testApiConnection() {
    try {
      console.log('=== Testing API Connection ===');
      const tokenInfo = getAccessToken();
      console.log('\nTesting with token:', tokenInfo.access_token.substring(0, 10) + '...');
      
      // Use instance_url from token response, falling back to config if not available
      const baseUrl = tokenInfo.instance_url || CONFIG.salesforceBaseUrl;
      const testUrl = baseUrl + CONFIG.apiVersionUrl;
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${tokenInfo.access_token}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      console.log('Sending test request to:', testUrl);
      const response = UrlFetchApp.fetch(testUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      console.log('Test response status:', responseCode);
      console.log('Test response body:', responseText);
      
      return {
        success: responseCode === 200,
        status: responseCode,
        response: responseText
      };
    } catch (error) {
      console.error('API test failed:', error.toString());
      return {
        success: false,
        error: error.toString()
      };
    }
}
