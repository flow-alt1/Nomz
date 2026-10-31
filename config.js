/* ============================================================
   NOMZ KITCHEN — site settings. This is the only file you need
   to edit when keys or links change.
   ============================================================ */
window.NOMZ = {
  // The Square Online store
  STORE_URL: "https://nomzkitchen.square.site/",

  // Where the final "Go to Square checkout" button sends the customer.
  // Square's cart icon lives on every store page, so the store home works.
  // If you find a direct cart URL on the Square site, paste it here.
  SQUARE_CART_URL: "https://nomzkitchen.square.site/",

  // Google Apps Script web app URL (ends in /exec). See README step 2.
  // Leave empty to run from menu.js only (no admin panel).
  API_URL: "https://script.google.com/macros/s/AKfycbzrIPlJmMK8ak11Y62XaNw2slUtVQ0JnKUBKCOnVq_NJIoeynEYOOnEOAPQkmWreSzZ6A/exec",

  // Web3Forms access key for the catering / contact form.
  WEB3FORMS_KEY: "6c4ea3d7-3394-4f60-a31e-8ff778031ef5",

  // How customers can receive orders at checkout.
  // ["Pickup"] = pickup only. ["Pickup", "Delivery"] = customer chooses
  // (Delivery asks for an address on the Square payment page).
  ORDER_TYPES: ["Pickup", "Delivery"],

  // Kitchen hours (24h clock), used for the "Open now" badge.
  // 0 = Sunday … 6 = Saturday. null = closed.
  TIMEZONE: "America/New_York",
  HOURS: {
    0: null,
    1: [11, 17],
    2: [11, 17],
    3: [11, 17],
    4: [11, 17],
    5: [11, 17],
    6: [11, 17]
  }
};
