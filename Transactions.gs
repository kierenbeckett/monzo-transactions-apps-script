var clientId = 'yourclientid';
var clientSecret = 'yourclientsecret';
var accountId = null;
var sheetName = 'Sheet1';

function updateTransactions() {
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spread.getSheetByName(sheetName);
  var scriptProperties = PropertiesService.getScriptProperties();

  if (scriptProperties.getProperty('refreshToken') == null) {
    Logger.info('Refresh token has not yet been set, follow the instructions to set up Monzo authentication');
    return;
  }

  if (new Date().getTime() > scriptProperties.getProperty('tokenExpiry')) {
    Logger.log('Token expired, refreshing');

    var payload = {
      'grant_type': 'refresh_token',
      'client_id': clientId,
      'client_secret': clientSecret,
      'refresh_token': scriptProperties.getProperty('refreshToken')
    };

    var url = 'https://api.monzo.com/oauth2/token';

    var options = {
      'method': 'post',
      'payload': payload
    };
    var data = JSON.parse(UrlFetchApp.fetch(url, options));

    scriptProperties.setProperty('token', data.access_token);
    scriptProperties.setProperty('refreshToken', data.refresh_token);
    scriptProperties.setProperty('tokenExpiry', new Date().getTime() + data.expires_in * 1000 - 10 * 60 * 1000);
  }

  do {
    if (accountId == null) {
      var url = 'https://api.monzo.com/accounts'

      var options = {
        'headers': {'Authorization' : 'Bearer ' + scriptProperties.getProperty('token')},
      };
      var data = JSON.parse(UrlFetchApp.fetch(url, options))['accounts'];

      Logger.info('Account ID has not yet been set, accounts are:')
      for (var i = 0; i<data.length; i++) {
        Logger.info(data[i]['description'] + ' = ' + data[i]['id']);
      }

      return;
    }

    var since = sheet.getRange(sheet.getDataRange().getLastRow(), 1, 1, 1).getValue()
    if (since == '') {
      var weekAgo = new Date();
      weekAgo.setDate(weekAgo.getDate() - 7);
      since = weekAgo.toISOString()
    }

    Logger.log('Fetching transactions since ' + since);

    var url = 'https://api.monzo.com/transactions'
    + '?account_id=' + accountId
    + '&expand[]=merchant'
    + '&since=' + encodeURIComponent(since);

    var options = {
      'headers': {'Authorization' : 'Bearer ' + scriptProperties.getProperty('token')},
    };

    var data = JSON.parse(UrlFetchApp.fetch(url, options))['transactions'];

    // TODO if there is a whole page of declined transactions we could miss later pages with valid transactions
    var processed = 0;
    for (var i = 0; i<data.length; i++) {
      var created = new Date(data[i]['created']);

      var name = data[i]['counterparty']['name'];
      var emoji = null;
      var address = null;
      if (data[i]['merchant'] != null) {
        name = data[i]['merchant']['name'];
        emoji = data[i]['merchant']['emoji'];
        address = data[i]['merchant']['address']['address'];
      }

      if (data[i]['decline_reason'] != null) {
        Logger.log('Transaction ' + data[i]['id'] + ' is declined ' + data[i]['decline_reason'] + ', skipping');
        continue;
      }

      row = [
        data[i]['id'],
        Utilities.formatDate(created, 'UTC', 'dd/MM/yyyy'),
        Utilities.formatDate(created, 'UTC', 'HH:mm:ss'),
        friendlyScheme(data[i]['scheme']),
        name,
        emoji,
        data[i]['category'][0].toUpperCase() + data[i]['category'].substr(1).replace('_', ' '),
        data[i]['amount'] / 100,
        data[i]['currency'],
        data[i]['local_amount'] / 100,
        data[i]['local_currency'],
        data[i]['notes'],
        address,
        null, // TODO Receipt?
        data[i]['description'],
        null, // TODO Category split?
        data[i]['amount'] <= 0 ? data[i]['amount'] / 100 : null,
        data[i]['amount'] > 0 ? data[i]['amount'] / 100 : null
      ];
      sheet.appendRow(row);
      processed++;
    }
  } while (processed > 0)
}

function friendlyScheme(scheme) {
  switch(scheme) {
    case 'mastercard':
      return 'Card payment';
    case 'bacs':
      return 'Direct Debit'
    case 'payport_faster_payments':
      return 'Faster payment';
    case 'p2p_payment':
      return 'P2P payment';
    case 'monzo_to_monzo':
      return 'Monzo-to-Monzo';
  }

  throw new Error('Unknown scheme ' + scheme);
}

function doGet(e) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var redirectUri = ScriptApp.getService().getUrl();

  if (e.parameter.code == null) {
    var nonce = '';
    var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for(var i=0; i < 200; i++) {
        nonce += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    scriptProperties.setProperty('nonce', nonce);

    var url = 'https://auth.monzo.com/'
      + '?redirect_uri=' + encodeURIComponent(redirectUri)
      + '&client_id=' + clientId
      + '&response_type=code'
      + '&state=' + nonce;

    return HtmlService.createHtmlOutput('<a href="' + url + '" target="_top">Click here</a>');
  }

  if(scriptProperties.getProperty('nonce') != e.parameter.state) {
    return ContentService.createTextOutput("Invalid state token");
  }

  var payload = {
    'grant_type': 'authorization_code',
    'client_id': clientId,
    'client_secret': clientSecret,
    'redirect_uri': redirectUri,
    'code': e.parameter.code
  };

  var url = 'https://api.monzo.com/oauth2/token';

  var options = {
    'method': 'post',
    'payload': payload
  };
  var data = JSON.parse(UrlFetchApp.fetch(url, options));

  scriptProperties.setProperty('token', data.access_token);
  scriptProperties.setProperty('refreshToken', data.refresh_token);
  scriptProperties.setProperty('tokenExpiry', new Date().getTime() + data.expires_in * 1000 - 10 * 60 * 1000);

  return ContentService.createTextOutput('Auth complete');
}

