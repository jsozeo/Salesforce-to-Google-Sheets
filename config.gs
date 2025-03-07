// OAuth 2.0 configuration for both PROD and ETI
const SALESFORCE_CONFIG = {
  ETI: {
    // OAuth endpoints
    accessTokenUrl: 'https://<replace>.my.salesforce.com/services/oauth2/token',
    
    // Salesforce API endpoints
    ENDPOINT: 'https://<replace>.lightning.force.com/services/data/v62.0/analytics/reports/<replace>',
    CLICK_LINK: 'https://<replace>.lightning.force.com/lightning/r/Report/<reportId>/view',
    apiVersionUrl: '/services/data/v62.0/',
    
    // Authentication
    clientId: '<replace>',
    clientSecret: '<replace>',
    sendCredentialsInBody: true
  },
  PROD: {
    // OAuth endpoints
    accessTokenUrl: 'https://<replace>.my.salesforce.com/services/oauth2/token',
    
    // Salesforce API endpoints
    ENDPOINT: 'https://<replace>.lightning.force.com/services/data/v62.0/analytics/reports/<replace>',
    CLICK_LINK: 'https://<replace>.lightning.force.com/lightning/r/Report/<prod_reportId>/view',
    apiVersionUrl: '/services/data/v62.0/',
    
    // Authentication
    clientId: '<replace>',
    clientSecret: '<replace>',
    sendCredentialsInBody: true
  }
}; 
