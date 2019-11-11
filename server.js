const sso = require('./node_modules/office-addin-sso/lib/server');
require('dotenv').config();

const ssoInstance = new sso.SSOService('./manifest.xml');
ssoInstance.startSsoService();