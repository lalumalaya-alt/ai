/**
 * Developed by Rameez Scripts
 * WhatsApp: https://wa.me/923224083545 (For Custom Projects)
 * YouTube: https://www.youtube.com/@rameezimdad (Subscribe for more!)
 */

/***************************************************
 * GLOBAL CONSTANTS
 ***************************************************/
const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// SHEET NAMES
const LOGIN_SHEET_NAME      = "Login";
const ROOMS_SHEET_NAME      = "Rooms";
const BOOKINGS_SHEET_NAME   = "Bookings";
const QUOTES_SHEET_NAME     = "Quotes";
const FINANCE_SHEET_NAME    = "Finance";
const INVOICES_SHEET_NAME   = "Invoices";
const SETTINGS_SHEET_NAME   = "Settings";
const BUDGETS_SHEET_NAME    = "Budgets";
const CATEGORIES_SHEET_NAME = "Categories";
const CUSTOMERS_SHEET_NAME  = "Customers";
const CHECKIN_SHEET_NAME    = "CheckIn";
const RESTAURANT_SHEET_NAME = "Restaurant";

// ROOMS sheet columns (0-based)
const ROOM_NO_COL          = 0;
const ROOM_TYPE_COL        = 1;
const ROOM_RATE_COL        = 2;
const ROOM_STATUS_COL      = 3;

// BOOKINGS sheet columns (0-based)
const TICKET_ID_COL        = 0;
const BOOKING_ROOM_NO_COL  = 1;
const GUEST_NAME_COL       = 2;
const PHONE_COL            = 3;
const EMAIL_COL            = 4;
const CITY_COL             = 5;
const MARITAL_STATUS_COL   = 6;
const OCCUPANCY_TYPE_COL   = 7;
const FAMILY_DETAILS_COL   = 8;
const CHECK_IN_COL         = 9;
const CHECK_OUT_COL        = 10;
const BOOKING_STATUS_COL   = 11;
const ROOM_RATE_BOOK_COL   = 12;
const DISCOUNT_COL         = 13;
const TAX_COL              = 14;
const PAYMENT_METHOD_COL   = 15;
const TOTAL_AMOUNT_COL     = 16;
const PAYMENT_STATUS_COL   = 17;  // "Unpaid", "Partial", "Paid"
const AMOUNT_PAID_COL      = 18;  // Numeric amount paid so far
const CHECKIN_TIME_COL     = 19;
const CHECKOUT_TIME_COL    = 20;
const FOOD_PLAN_COL        = 21;
const ADVANCE_PAID_COL     = 22;
const NUM_ROOMS_COL        = 23;
const LINKED_CHECKIN_COL   = 24;

// LOGIN sheet columns (0-based)
const LOGIN_USERNAME_COL   = 0;
const LOGIN_PASSWORD_COL   = 1;
const LOGIN_ROLE_COL       = 2;
const LOGIN_OTP_COL        = 3;
const LOGIN_OTP_EXPIRY_COL = 4;

// QUOTES sheet columns (0-based)
const QUOTE_ID_COL              = 0;
const QUOTE_GUEST_NAME_COL      = 1;
const QUOTE_PHONE_COL           = 2;
const QUOTE_EMAIL_COL           = 3;
const QUOTE_CREATED_COL         = 4;
const QUOTE_VALID_COL           = 5;
const QUOTE_STATUS_COL          = 6;
const QUOTE_ITEMS_COL           = 7;
const QUOTE_SUBTOTAL_COL        = 8;
const QUOTE_TAX_COL             = 9;
const QUOTE_DISCOUNT_COL        = 10;
const QUOTE_TOTAL_COL           = 11;
const QUOTE_NOTES_COL           = 12;
const QUOTE_CREATED_BY_COL      = 13;
const QUOTE_CURRENCY_COL        = 14;
const QUOTE_GST_ENABLED_COL     = 15;
const QUOTE_GST_PERCENT_COL     = 16;
const QUOTE_GST_AMOUNT_COL      = 17;
const QUOTE_GREENTAX_ENABLED_COL= 18;
const QUOTE_GREENTAX_RATE_COL   = 19;
const QUOTE_GREENTAX_PAX_COL    = 20;
const QUOTE_GREENTAX_NIGHTS_COL = 21;
const QUOTE_GREENTAX_AMOUNT_COL = 22;
const QUOTE_CUSTOMER_TIN_COL    = 23;
const QUOTE_CONVERTED_COL       = 24;
const QUOTE_PDF_LINK_COL        = 25;

// FINANCE sheet columns (0-based)
const FIN_ID_COL           = 0;
const FIN_DATE_COL         = 1;
const FIN_TYPE_COL         = 2;
const FIN_DESC_COL         = 3;
const FIN_SHOP_COL         = 4;
const FIN_AMOUNT_COL       = 5;
const FIN_BALANCE_COL      = 6;
const FIN_ENTERED_BY_COL   = 7;
const FIN_CREATED_AT_COL   = 8;
const FIN_CATEGORY_COL     = 9;
const FIN_CURRENCY_COL     = 10;
const FIN_LINKED_INV_COL   = 11;

// INVOICES sheet columns (0-based)
const INV_ID_COL              = 0;
const INV_GUEST_NAME_COL      = 1;
const INV_PHONE_COL           = 2;
const INV_EMAIL_COL           = 3;
const INV_CUSTOMER_TIN_COL    = 4;
const INV_CURRENCY_COL        = 5;
const INV_CREATED_DATE_COL    = 6;
const INV_DUE_DATE_COL        = 7;
const INV_STATUS_COL          = 8;
const INV_ITEMS_COL           = 9;
const INV_SUBTOTAL_COL        = 10;
const INV_GST_ENABLED_COL     = 11;
const INV_GST_PERCENT_COL     = 12;
const INV_GST_AMOUNT_COL      = 13;
const INV_GREENTAX_ENABLED_COL= 14;
const INV_GREENTAX_RATE_COL   = 15;
const INV_GREENTAX_PAX_COL    = 16;
const INV_GREENTAX_NIGHTS_COL = 17;
const INV_GREENTAX_AMOUNT_COL = 18;
const INV_DISCOUNT_COL        = 19;
const INV_TOTAL_COL           = 20;
const INV_NOTES_COL           = 21;
const INV_SOURCE_QUOTE_COL    = 22;
const INV_PDF_LINK_COL        = 23;
const INV_CREATED_BY_COL      = 24;
const INV_UPDATED_AT_COL      = 25;

// SETTINGS sheet columns (0-based, single data row at row 2)
const SET_HOTEL_NAME_COL       = 0;
const SET_HOTEL_ADDRESS_COL    = 1;
const SET_HOTEL_PHONE_COL      = 2;
const SET_HOTEL_EMAIL_COL      = 3;
const SET_HOTEL_TIN_COL        = 4;
const SET_LOGO_FILE_ID_COL     = 5;
const SET_LOGO_URL_COL         = 6;
const SET_DEFAULT_CURRENCY_COL = 7;
const SET_GST_DEFAULT_COL      = 8;
const SET_GREENTAX_DEFAULT_COL = 9;
const SET_NEXT_INVOICE_COL     = 10;
const SET_NEXT_QUOTE_COL       = 11;
const SET_PDF_FOLDER_ID_COL    = 12;
const SET_LOGO_FOLDER_ID_COL   = 13;
const SET_NEXT_CHECKIN_COL  = 14;
const SET_NEXT_BILL_COL     = 15;

// BUDGETS sheet columns (0-based)
const BDG_ID_COL           = 0;
const BDG_MONTH_COL        = 1;
const BDG_YEAR_COL         = 2;
const BDG_AMOUNT_COL       = 3;
const BDG_SPENT_COL        = 4;
const BDG_REMAINING_COL    = 5;
const BDG_SET_BY_COL       = 6;
const BDG_CREATED_AT_COL   = 7;
const BDG_UPDATED_AT_COL   = 8;

// CATEGORIES sheet columns (0-based)
const CAT_ID_COL           = 0;
const CAT_NAME_COL         = 1;
const CAT_TYPE_COL         = 2;
const CAT_IS_DEFAULT_COL   = 3;
const CAT_CREATED_BY_COL   = 4;
const CAT_CREATED_AT_COL   = 5;

// CHECKIN sheet columns (0-based)
const CI_ID_COL             = 0;
const CI_LINKED_TICKET_COL  = 1;
const CI_GUEST_NAME_COL     = 2;
const CI_COMPANY_COL        = 3;
const CI_GST_NUMBER_COL     = 4;
const CI_IDENTITY_COL       = 5;
const CI_MOBILE_COL         = 6;
const CI_EMAIL_COL          = 7;
const CI_ADDRESS_COL        = 8;
const CI_PURPOSE_COL        = 9;
const CI_CHECKIN_DATE_COL   = 10;
const CI_CHECKIN_TIME_COL   = 11;
const CI_CHECKOUT_DATE_COL  = 12;
const CI_CHECKOUT_TIME_COL  = 13;
const CI_ROOM_NUMBERS_COL   = 14;
const CI_ROOM_TYPES_COL     = 15;
const CI_NUM_ROOMS_COL      = 16;
const CI_PAX_COL            = 17;
const CI_ADVANCE_PAID_COL   = 18;
const CI_EXTRA_PERSON_COL   = 19;
const CI_FOOD_PLAN_COL      = 20;
const CI_GST_TYPE_COL       = 21;
const CI_FIX_RENT_COL       = 22;
const CI_FIX_RENT_AMT_COL   = 23;
const CI_BILL_TO_COL        = 24;
const CI_DISCOUNT_COL       = 25;
const CI_STATUS_COL         = 26;
const CI_CREATED_AT_COL     = 27;

// RESTAURANT sheet columns (0-based)
const REST_ORDER_ID_COL     = 0;
const REST_ROOM_NO_COL      = 1;
const REST_CHECKIN_ID_COL   = 2;
const REST_ORDER_DATE_COL   = 3;
const REST_CATEGORY_COL     = 4;
const REST_DESC_COL         = 5;
const REST_AMOUNT_COL       = 6;
const REST_STATUS_COL       = 7;
const REST_CREATED_AT_COL   = 8;

// CUSTOMERS sheet columns (0-based)
const CUST_ID_COL           = 0;
const CUST_NAME_COL         = 1;
const CUST_PHONE_COL        = 2;
const CUST_EMAIL_COL        = 3;
const CUST_CITY_COL         = 4;
const CUST_MARITAL_COL      = 5;
const CUST_NOTES_COL        = 6;
const CUST_CREATED_AT_COL   = 7;
const CUST_LINKED_USER_COL  = 8;
