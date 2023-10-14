/*
With this Google Sheet script, we can create customers and subscriptions in Stripe from a Google Sheet.
The script will also email the customer a checkout link to complete the subscription setup.
The script will also update the subscription status in the Google Sheet.


/**
 * Sets the API key in the script properties.
*/
function setApiKeyToScriptProperties() {
  let apiKey = settings.getSetting(SETTINGS_API_KEY_LABEL);
  if (!apiKey) {
    //notify user that the API key is not set
    SpreadsheetApp.getUi().alert('API Key not set');
    return;
  }
  settings.setSettingInScriptProperties(SETTINGS_API_KEY_LABEL, apiKey);
}

function onOpen() {
  settings.init();

  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Stripe')
    .addItem("Create Subscriptions", 'subscribeCustomers')
    .addToUi();
}

/**
 * Fetches customer data from the active Google Sheet.
 * @returns {Array} List of customer objects.
 */
function fetchCustomersFromSheet() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRow = sheet.getLastRow();

  // Assuming the data starts from the second row (considering the first row has headers)
  let dataRange = sheet.getRange(2, 1, lastRow - 1, 10);
  let customers = dataRange.getValues();

  let customerList = [];

  for (const element of customers) {
    let customer = {
      "name": element[0],
      "email": element[1],
      "address": element[2],
      "country": element[3],
      "city": element[4],
      "postalCode": element[5],
      "months": element[6],
      "trialPeriod": element[7],
      "amount": element[8],
      "subscriptionStatus": element[8]
    };

    customerList.push(customer);
  }

  return customerList;
}

/**
 * Fetches or creates a price for a product in Stripe.
 * @param {number} unit_amount - Amount for the product.
 * @param {string} product - Stripe product ID.
 * @param {string} currency - Currency code (e.g., 'EUR').
 * @returns {string|null} Stripe price ID or null if not found.
 */
function getPrice(unit_amount, product, currency) {

  let prices = stripeClient.fetchProductPrices(product);

  for (const element of prices.data) {
    if (element.unit_amount == unit_amount && element.currency.toLowerCase() == currency.toLowerCase()) {
      // If a price with the specified unit_amount exists, return its ID
      return element.id;
    }
  }


  return null;
}

/**
 * Sets a specific field in the Google Sheet based on the provided email.
 * @param {string} email - Customer's email.
 * @param {number} column - Column number to set the value.
 * @param {string} value - Value to set in the specified column.
 */
function setField(email, column, value) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  let lastRow = sheet.getLastRow();

  // The data starts from the second row (considering the first row has headers)
  let emailRange = sheet.getRange(2, 2, lastRow - 1, 1);
  let emails = emailRange.getValues();

  for (let i = 0; i < emails.length; i++) {
    if (emails[i][0] === email) {
      sheet.getRange(i + 2, column).setValue(value);
      break;
    }
  }
}

/**
 * Main function to initiate the subscription process for customers.
 */
function subscribeCustomers() {
  try {
    // Start by checking if the required settings exist and notify the user if they don't
    let apiKey = settings.getSettingFromScriptProperties(SETTINGS_API_KEY_LABEL);
    let selectedProductId = settings.getSetting(SETTINGS_SELECTED_PRODUCT_LABEL);
    let defaultTrialPeriod = settings.getSetting(SETTINGS_DEFAULT_TRIAL_PERIOD_LABEL);

    if (!apiKey) {
      throw new Error('API Key not found in script properties.');
    }

    if (!selectedProductId) {
      throw new Error('Selected Product not found in settings.');
    }

    if (defaultTrialPeriod == null) {
      throw new Error('Default Trial Period not found in settings.');
    }

    // Get Customers from sheet
    let sheetCustomers = fetchCustomersFromSheet();

    for (const element of sheetCustomers) {
      let fc = element;
      if (fc.subscriptionStatus != 'Subscribed') {
        subscribeCustomer(fc, selectedProductId, defaultTrialPeriod);
      }
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}


/**
 * Handles the subscription process for a single customer.
 * @param {Object} fc - Customer object.
 * @param {string} productId - Stripe product ID.
 * @param {number} defaultTrialPeriod - Default trial period in days.
 */
function subscribeCustomer(fc, productId, defaultTrialPeriod) {
  try {
    // Create Customer
    let newStripeCustomer = stripeClient.createCustomer(fc.name, fc.email, fc.address, fc.country, fc.city, fc.postalCode);
    if (!newStripeCustomer) {
      throw new Error('Failed to create a new Stripe customer.');
    }

    let amount = fc.amount * 100;
    let product = stripeClient.fetchProduct(productId);

    if (product == null) {
      throw new Error('Product not found in Stripe.');
    }

    // Check if the price exists, if not create it
    let priceId = getPrice(amount, product.id, CURRENCY);
    if (priceId == null) {
      throw new Error('Price not found for the selected product in Stripe.');
    }

    let trialPeriod = (fc.trialPeriod != null && fc.trialPeriod != "") ? parseInt(fc.trialPeriod) : defaultTrialPeriod;
    let subs = stripeClient.createSubscription(newStripeCustomer.id, priceId, fc.months, trialPeriod);

    if (!subs) {
      throw new Error('Failed to create a subscription in Stripe.');
    }

    let session = stripeClient.createCheckoutSession(priceId, newStripeCustomer.id, "card", "https://apptivasoftware.com");
    if (!session) {
      throw new Error('Failed to create a checkout session in Stripe.');
    }

    sendCheckoutLink(newStripeCustomer, session);
    setField(fc.email, 10, 'CREATED');
    setField(fc.email, 11, subs.id);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error for customer ${fc.email}: ${error.message}`);
  }
}


/**
 * Sends a checkout link to the customer via email.
 * @param {Object} customer - Stripe customer object.
 * @param {Object} session - Stripe checkout session object.
 */
function sendCheckoutLink(customer, session) {
  let checkoutUrl = session.url;

  // Send an email to the customer with the checkout URL
  let subject = "Complete Your Subscription Setup";
  let body = `
        Dear ${customer.name},<br><br>
        Please complete your subscription setup by <a href="${checkoutUrl}">clicking here</a>.<br><br>
        Regards,<br>
        Your Company Name
    `;

  MailApp.sendEmail({
    to: customer.email,
    subject: subject,
    body: "",
    htmlBody: body
  });
}