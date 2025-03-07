// Main function to execute the script
function main() {
    try {
      console.log('=== Starting script execution ===');
      
      // Test both environments
      console.log('\nTesting API connections...');
      const testResults = testAllConnections();
      if (!testResults.ETI.success && !testResults.PROD.success) {
        throw new Error('API connection tests failed for both environments');
      }
      console.log('At least one environment connection test successful');
      
      return testResults;
    } catch (error) {
      console.error('=== Script execution failed ===');
      console.error('Error:', error.toString());
      throw error;
    }
}

// Function to get OAuth 2.0 access token using Client Credentials grant
function getAccessToken(environment) {
    console.log(`Starting OAuth token acquisition for ${environment}...`);
    
    const config = SALESFORCE_CONFIG[environment];
    if (!config) {
        throw new Error(`No configuration found for environment: ${environment}`);
    }
    
    const payload = {
      grant_type: 'client_credentials'
    };
    
    let options = {
      method: 'post',
      muteHttpExceptions: true
    };
    
    // Add client authentication based on configuration
    if (config.sendCredentialsInBody) {
      console.log('Using body authentication method');
      payload.client_id = config.clientId;
      payload.client_secret = config.clientSecret;
      options.payload = payload;
      console.log('Client credentials added to request body');
    } else {
      console.log('Using header authentication method');
      // Send credentials in Authorization header
      const credentials = Utilities.base64Encode(config.clientId + ':' + config.clientSecret);
      options.headers = {
        'Authorization': 'Basic ' + credentials
      };
      options.payload = payload;
      console.log('Client credentials added to Authorization header');
    }
    
    try {
      console.log('Sending token request to:', config.accessTokenUrl);
      const response = UrlFetchApp.fetch(config.accessTokenUrl, options);
      console.log('Token request response received');
      
      const responseText = response.getContentText();
      console.log('Token response text received');
      
      const result = JSON.parse(responseText);
      console.log(`\n=== TOKEN INFORMATION FOR ${environment} ===`);
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
      console.error(`Error getting access token for ${environment}:`, error.toString());
      throw error;
    }
}

// Function to fetch Salesforce report data
function getSalesforceReport(environment) {
    try {
      console.log(`Starting Salesforce report retrieval for ${environment}...`);
      const tokenInfo = getAccessToken(environment);
      console.log(`Successfully obtained access token for ${environment}:`, tokenInfo.access_token.substring(0, 10) + '...');
      
      const config = SALESFORCE_CONFIG[environment];
      if (!config) {
          throw new Error(`No configuration found for environment: ${environment}`);
      }
      
      console.log('Preparing request to:', config.ENDPOINT);
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${tokenInfo.access_token}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      console.log('Request headers:', JSON.stringify(options.headers));
      console.log(`Sending request to Salesforce API for ${environment}...`);
      
      const response = UrlFetchApp.fetch(config.ENDPOINT, options);
      const responseCode = response.getResponseCode();
      console.log('Response status:', responseCode);
      
      if (responseCode !== 200) {
        const errorText = response.getContentText();
        console.error('Error response:', errorText);
        throw new Error(`API request failed with status ${responseCode}: ${errorText}`);
      }
      
      const responseText = response.getContentText();
      console.log('Raw response length:', responseText.length);
      
      const result = JSON.parse(responseText);
      console.log('Response successfully parsed as JSON');
      
      // Log the report details
      console.log(`\nReport Details for ${environment}:`);
      console.log('Report Name:', result.attributes?.reportName || 'Not found');
      console.log('Total Records:', result.factMap?.['T!T']?.aggregates?.[0]?.value || 'Not found');
      
      return result;
    } catch (error) {
      console.error(`Error fetching Salesforce report for ${environment}:`, error.toString());
      if (error.stack) {
        console.error('Stack trace:', error.stack);
      }
      throw error;
    }
}

// Function to test API connection for a specific environment
function testApiConnection(environment) {
    try {
      console.log(`=== Testing API Connection for ${environment} ===`);
      const tokenInfo = getAccessToken(environment);
      console.log('\nTesting with token:', tokenInfo.access_token.substring(0, 10) + '...');
      
      const config = SALESFORCE_CONFIG[environment];
      if (!config) {
          throw new Error(`No configuration found for environment: ${environment}`);
      }
      
      // Use instance_url from token response or extract domain from ENDPOINT
      const baseUrl = tokenInfo.instance_url || new URL(config.ENDPOINT).origin;
      const testUrl = baseUrl + config.apiVersionUrl;
      
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
      
      console.log(`Test response status for ${environment}:`, responseCode);
      console.log('Test response body:', responseText);
      
      return {
        success: responseCode === 200,
        status: responseCode,
        response: responseText
      };
    } catch (error) {
      console.error(`API test failed for ${environment}:`, error.toString());
      return {
        success: false,
        error: error.toString()
      };
    }
}

// Main function to test both environments
function testAllConnections() {
    console.log('=== Testing All Environment Connections ===');
    
    const results = {
        ETI: null,
        PROD: null
    };
    
    try {
        results.ETI = testApiConnection('ETI');
        console.log('\nETI Connection Test Result:', results.ETI.success ? 'SUCCESS' : 'FAILED');
    } catch (error) {
        console.error('ETI Connection Test Error:', error.toString());
        results.ETI = { success: false, error: error.toString() };
    }
    
    try {
        results.PROD = testApiConnection('PROD');
        console.log('\nPROD Connection Test Result:', results.PROD.success ? 'SUCCESS' : 'FAILED');
    } catch (error) {
        console.error('PROD Connection Test Error:', error.toString());
        results.PROD = { success: false, error: error.toString() };
    }
    
    return results;
}
