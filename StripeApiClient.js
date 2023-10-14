/**
 * Stripe API Client for handling various Stripe operations.
 */
class StripeApiClient {
    constructor(apiKey) {
        this.apiKey = apiKey;
    }

    /**
     * Private method to make requests to the Stripe API.
     * @param {string} endpoint - The Stripe API endpoint.
     * @param {string} method - The HTTP method (e.g., 'post', 'get').
     * @param {object} data - The data payload for the request.
     * @returns {object} - The parsed JSON response from the Stripe API.
     */
    _makeRequest(endpoint, method, data) {
        let headers = {
            "Authorization": "Bearer " + this.apiKey,
            "Content-Type": "application/x-www-form-urlencoded"
        };

        let options = {
            "method": method,
            "headers": headers,
            "payload": data
        };

        let response = UrlFetchApp.fetch('https://api.stripe.com/v1/' + endpoint, options);
        return JSON.parse(response.getContentText());
    }

    /**
   * Create a new customer in Stripe.
   * @param {string} name - The customer's name.
   * @param {string} email - The customer's email address.
   * @param {string} address - The customer's address.
   * @param {string} country - The customer's country.
   * @param {string} city - The customer's city.
   * @param {string} postalCode - The customer's postal code.
   * @returns {object} - The created customer object.
   */
    createCustomer(name, email, address, country, city, postalCode) {
        let customerData = {
            "name": name,
            "email": email,
            "address[line1]": address,
            "address[city]": city,
            "address[postal_code]": postalCode,
            "address[country]": country
        };

        return this._makeRequest('customers', 'post', customerData);
    }

    /**
   * Create a subscription for a customer.
   * @param {string} customerId - The ID of the customer.
   * @param {string} priceId - The ID of the price.
   * @param {number} duration - The duration of the subscription in months.
   * @param {number} trialDays - The trial period in days.
   * @returns {object} - The created subscription object.
   */
    createSubscription(customerId, priceId, duration, trialDays) {
        let currentDate = new Date();
        currentDate.setMonth(currentDate.getMonth() + parseInt(duration));
        let cancelAtTimestamp = Math.floor(currentDate.getTime() / 1000);

        let subscriptionData = {
            "customer": customerId,
            "items[0][price]": priceId,
            "cancel_at": cancelAtTimestamp.toString(),
          
        };

        if(trialDays > 0) {
            let trialEndDate = new Date();
            trialEndDate.setDate(trialEndDate.getDate() + trialDays);
            let trialEndTimestamp = Math.floor(trialEndDate.getTime() / 1000);

            subscriptionData["trial_end"] = trialEndTimestamp.toString();
        }

        return this._makeRequest('subscriptions', 'post', subscriptionData);
    }

    /**
     * Create a checkout session for a customer.
     * @param {string} priceId - The ID of the price.
     * @param {string} customerId - The ID of the customer.
     * @param {string} [success_url="https://your-success-url.com"] - The URL to redirect to upon successful checkout.
     * @param {string} [cancel_url="https://your-cancel-url.com"] - The URL to redirect to upon checkout cancellation.
     * @returns {object} - The created checkout session object.
     */
    createCheckoutSession(priceId, customerId, payment_method_type = "card", success_url = "https://your-success-url.com", cancel_url = "https://your-cancel-url.com") {
        let sessionData = {
            "customer": customerId,
            "payment_method_types[]": payment_method_type,
            "line_items[0][price]": priceId,
            "line_items[0][quantity]": "1",
            "mode": "subscription",
            "success_url": success_url,
            "cancel_url": cancel_url
        };

        return this._makeRequest('checkout/sessions', 'post', sessionData);
    }

    /**
     * Fetch a list of customers from Stripe.
     * @returns {object} - The list of customers.
     */
    fetchCustomers() {
        return this._makeRequest('customers', 'get', {});
    }

    /**
     * Fetch a list of products from Stripe.
     * @returns {object} - The list of products.
     */
    fetchProducts() {
        return this._makeRequest('products', 'get', {});
    }

    fetchProduct(productId) {
        return this._makeRequest('products/' + productId, 'get', {});
    }

    /**
     * Fetch a list of prices for a specific product from Stripe.
     * @param {string} product - The ID of the product.
     * @returns {object} - The list of prices for the specified product.
     */
    fetchProductPrices(product) {
        return this._makeRequest('prices?product=' + product, 'get', {});
    }


    /**
    * Create a new price for a product in Stripe.
    * @param {number} unit_amount - The amount for the price.
    * @param {string} product - The ID of the product.
    * @param {string} currency - The currency for the price.
    * @returns {object} - The created price object.
    */
    createProductPrice(unit_amount, product, currency) {
        let priceData = {
            "product": product,
            "unit_amount": unit_amount.toString(),
            "currency": currency,
            "recurring[interval]": "month"
        };

        return this._makeRequest('prices', 'post', priceData);
    }
}