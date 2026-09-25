
/****************************************************
 * SEEMAN FOODS - GOOGLE SHEETS BACKEND
 * Complete Stable Version
 ****************************************************/

const SPREADSHEET_ID =
  '1tUldiwodSrnahoCwoOsD_czcYDsnuJXrCAwPDpHEExs';

const TZ = Session.getScriptTimeZone() || 'Asia/Colombo';

/* ==================================================
   SHEET HEADERS
================================================== */

const HEADERS = {

  users: [
    'user_id',
    'full_name',
    'username',
    'password',
    'role',
    'email',
    'contact',
    'status'
  ],

  products: [
    'product_id',
    'product_code',
    'product_name',
    'unit_price'
  ],

  customers: [
    'customer_id',
    'customer_code',
    'customer_name',
    'contact_number',
    'email',
    'address',
    'registered_date'
  ],

  sales: [
    'sales_id',
    'sales_code',
    'customer_id',
    'product_id',
    'quantity',
    'unit_price',
    'service_charge_rate',
    'commission_amount',
    'subtotal',
    'service_charge_amount',
    'total_amount',
    'payment_status',
    'sale_date',
    'created_by'
  ],

  costs: [
    'cost_id',
    'cost_code',
    'product_name',
    'category',
    'quantity',
    'unit_price',
    'total_price',
    'cost_date',
    'payment_status',
    'created_by'
  ],

  orders: [
    'order_id',
    'order_code',
    'customer_id',
    'order_date',
    'quantity',
    'order_status',
    'estimated_delivery_date',
    'delivery_address',
    'created_by'
  ],

  payments: [
    'payment_id',
    'payment_code',
    'customer_id',
    'sales_id',
    'payment_date',
    'payment_amount',
    'payment_method',
    'transaction_ref',
    'payment_status',
    'created_by'
  ],

  receipts: [
    'receipt_id',
    'receipt_code',
    'payment_id',
    'sales_id',
    'customer_id',
    'amount_paid',
    'total_sale_amount',
    'total_paid',
    'balance_due',
    'payment_method',
    'payment_date',
    'issued_by',
    'issued_at',
    'status'
  ],

  reports: [
    'report_id',
    'report_code',
    'report_month',
    'report_year',
    'total_gross_sales',
    'total_service_charges',
    'total_costs',
    'total_commissions',
    'net_profit',
    'generated_by',
    'generated_at'
  ],

  login_logs: [
    'log_id',
    'user_id',
    'username',
    'full_name',
    'role',
    'login_time',
    'action',
    'status',
    'message'
  ]
};


/* ==================================================
   BASIC HELPERS
================================================== */

function getSS() {

  const active = SpreadsheetApp.getActiveSpreadsheet();

  if (active) {
    return active;
  }

  if (!SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}


function sheet(name) {

  if (!HEADERS[name]) {
    throw new Error('Unsupported sheet: ' + name);
  }

  const ss = getSS();

  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  ensureHeaders(sh, name);

  return sh;
}


function ensureHeaders(sh, name) {

  const headers = HEADERS[name];

  if (sh.getLastRow() === 0) {

    sh
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    sh.setFrozenRows(1);

    return;
  }

  const range =
    sh.getRange(1, 1, 1, headers.length);

  const current = range.getValues()[0];

  let changed = false;

  headers.forEach(function(header, index) {

    if (current[index] !== header) {

      current[index] = header;
      changed = true;

    }

  });

  if (changed) {
    range.setValues([current]);
  }

  sh.setFrozenRows(1);
}


/* ==================================================
   RESPONSE HELPERS
================================================== */

function json(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

}


function jsonp(data, callback) {

  const cb = String(callback || '')
    .replace(/[^a-zA-Z0-9_$.]/g, '');

  if (!cb) {
    return json(data);
  }

  return ContentService
    .createTextOutput(
      cb + '(' + JSON.stringify(data) + ');'
    )
    .setMimeType(ContentService.MimeType.JAVASCRIPT);

}


/* ==================================================
   GET API
================================================== */

function doGet(e) {

  try {

    const params =
      e && e.parameter ? e.parameter : {};

    const action =
      params.action || 'ping';

    let result;

    switch (action) {

      case 'getData':

        result = {
          success: true,
          data: getAllData()
        };

        break;


      case 'getDashboard':

        result = {
          success: true,
          data: getDashboard()
        };

        break;


      case 'setup':

        result = {
          success: true,
          message: setupSheets()
        };

        break;


      case 'ping':

        result = {
          success: true,
          message: 'Seeman Foods API is running'
        };

        break;


      default:

        result = {
          success: true,
          message: 'Seeman Foods API is running'
        };

    }

    return jsonp(
      result,
      params.callback
    );

  } catch (error) {

    return jsonp(
      {
        success: false,
        message: error.message
      },
      e && e.parameter
        ? e.parameter.callback
        : ''
    );

  }

}


/* ==================================================
   POST API
================================================== */

function doPost(e) {

  const lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(15000);

    const body =
      e &&
      e.postData &&
      e.postData.contents
        ? e.postData.contents
        : '{}';

    const d = JSON.parse(body);

    const action = d.action;

    if (!action) {
      throw new Error('Action is required.');
    }

    let result;

    switch (action) {

      /* SALES */

      case 'addSale':
        result = addSale(d);
        break;

      case 'updateSale':
        result = updateSale(d);
        break;

      case 'deleteSale':
        result = deleteRecord(
          'sales',
          'sales_id',
          d.sales_id
        );
        break;


      /* COSTS */

      case 'addCost':
        result = addCost(d);
        break;

      case 'updateCost':
        result = updateCost(d);
        break;

      case 'deleteCost':
        result = deleteRecord(
          'costs',
          'cost_id',
          d.cost_id
        );
        break;


      /* CUSTOMERS */

      case 'addCustomer':
        result = addCustomer(d);
        break;

      case 'updateCustomer':
        result = updateCustomer(d);
        break;

      case 'deleteCustomer':
        result = deleteRecord(
          'customers',
          'customer_id',
          d.customer_id
        );
        break;


      /* ORDERS */

      case 'addOrder':
        result = addOrder(d);
        break;

      case 'updateOrder':
        result = updateOrder(d);
        break;

      case 'deleteOrder':
        result = deleteRecord(
          'orders',
          'order_id',
          d.order_id
        );
        break;


      /* PAYMENTS */

      case 'addPayment':
        result = addPayment(d);
        break;

      case 'updatePayment':
        result = updatePayment(d);
        break;

      case 'deletePayment':
        result = deleteRecord(
          'payments',
          'payment_id',
          d.payment_id
        );
        break;


      /* RECEIPTS */

      case 'addReceipt':
        result = addReceipt(d);
        break;

      case 'updateReceipt':
        result = updateReceipt(d);
        break;

      case 'deleteReceipt':
        result = deleteRecord(
          'receipts',
          'receipt_id',
          d.receipt_id
        );
        break;


      /* REPORTS */

      case 'addReport':
        result = addReport(d);
        break;

      case 'deleteReport':
        result = deleteRecord(
          'reports',
          'report_id',
          d.report_id
        );
        break;


      /* USERS */

      case 'addUser':
        result = addUser(d);
        break;

      case 'updateUser':
        result = updateUser(d);
        break;

      case 'deleteUser':
        result = deleteRecord(
          'users',
          'user_id',
          d.user_id
        );
        break;


      /* PRODUCTS */

      case 'addProduct':
        result = addProduct(d);
        break;

      case 'updateProduct':
        result = updateProduct(d);
        break;

      case 'deleteProduct':
        result = deleteRecord(
          'products',
          'product_id',
          d.product_id
        );
        break;


      /* LOGIN / USER ACTIVITY LOG */

      case 'addLoginLog':
        result = addLoginLog(d);
        break;


      /* FULL TABLE SYNC */

      case 'replaceData':

        result = replaceData(
          d.table,
          d.rows || []
        );

        break;


      /* SETUP */

      case 'setupSheets':

        result = setupSheets();

        break;


      default:

        throw new Error(
          'Invalid action: ' + action
        );

    }

    return json({
      success: true,
      data: result
    });

  } catch (error) {

    return json({
      success: false,
      message: error.message,
      stack: error.stack
    });

  } finally {

    try {
      lock.releaseLock();
    } catch (ignore) {}

  }

}


/* ==================================================
   READ DATA
================================================== */

function read(name) {

  const sh = sheet(name);

  const lastRow =
    sh.getLastRow();

  const headers =
    HEADERS[name];

  if (lastRow < 2) {
    return [];
  }

  const values =
    sh
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();

  return values

    .filter(function(row) {

      return row.some(function(value) {

        return value !== '' &&
               value !== null;

      });

    })

    .map(function(row) {

      const obj = {};

      headers.forEach(function(header, index) {

        let value = row[index];

        if (value instanceof Date) {

          value =
            Utilities.formatDate(
              value,
              TZ,
              'yyyy-MM-dd'
            );

        }

        obj[header] = value;

      });

      return obj;

    });

}


/* ==================================================
   GET ALL DATA
================================================== */

function getAllData() {

  const result = {};

  Object.keys(HEADERS).forEach(function(name) {

    result[name] = read(name);

  });

  return result;

}


/* ==================================================
   ID / CODE HELPERS
================================================== */

function nextId(name) {

  const field =
    HEADERS[name][0];

  const rows =
    read(name);

  let maxId = 0;

  rows.forEach(function(row) {

    const id =
      Number(row[field]) || 0;

    if (id > maxId) {
      maxId = id;
    }

  });

  return maxId + 1;

}


function code(prefix, id) {

  return (
    prefix +
    '-' +
    new Date().getFullYear() +
    '-' +
    String(id).padStart(6, '0')
  );

}


/* ==================================================
   ROW HELPERS
================================================== */

function findRow(
  name,
  idField,
  id
) {

  const sh =
    sheet(name);

  const column =
    HEADERS[name].indexOf(idField) + 1;

  if (
    column < 1 ||
    sh.getLastRow() < 2
  ) {
    return null;
  }

  const values =
    sh
      .getRange(
        2,
        column,
        sh.getLastRow() - 1,
        1
      )
      .getValues();

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(id)
    ) {

      return i + 2;

    }

  }

  return null;

}


function rowObject(
  name,
  row
) {

  const sh =
    sheet(name);

  const headers =
    HEADERS[name];

  const values =
    sh
      .getRange(
        row,
        1,
        1,
        headers.length
      )
      .getValues()[0];

  const obj = {};

  headers.forEach(function(header, index) {

    obj[header] =
      values[index];

  });

  return obj;

}


function setByHeader(
  name,
  row,
  data
) {

  const sh =
    sheet(name);

  const headers =
    HEADERS[name];

  headers.forEach(function(header, index) {

    if (
      Object.prototype.hasOwnProperty.call(
        data,
        header
      )
    ) {

      sh
        .getRange(
          row,
          index + 1
        )
        .setValue(
          data[header]
        );

    }

  });

}


function appendByHeader(
  name,
  data
) {

  const headers =
    HEADERS[name];

  const row =
    headers.map(function(header) {

      return Object.prototype.hasOwnProperty.call(
        data,
        header
      )
        ? data[header]
        : '';

    });

  sheet(name).appendRow(row);

}


/* ==================================================
   NUMBER / DATE HELPERS
================================================== */

function num(value) {

  const n =
    Number(value);

  return isFinite(n)
    ? n
    : 0;

}


function dateValue(value) {

  if (!value) {
    return new Date();
  }

  if (value instanceof Date) {
    return value;
  }

  const date =
    new Date(value);

  if (isNaN(date.getTime())) {
    return new Date();
  }

  return date;

}


/* ==================================================
   SALES
================================================== */

function addSale(d) {

  const id =
    nextId('sales');

  const quantity =
    num(d.quantity);

  const unitPrice =
    num(d.unit_price);

  const serviceRate =
    num(d.service_charge_rate);

  const commission =
    num(d.commission_amount);

  const subtotal =
    quantity * unitPrice;

  const serviceCharge =
    subtotal *
    serviceRate /
    100;

  const total =
    subtotal -
    serviceCharge -
    commission;

  const record = {

    sales_id:
      id,

    sales_code:
      d.sales_code ||
      code('SAL', id),

    customer_id:
      d.customer_id || '',

    product_id:
      d.product_id || '',

    quantity:
      quantity,

    unit_price:
      unitPrice,

    service_charge_rate:
      serviceRate,

    commission_amount:
      commission,

    subtotal:
      subtotal,

    service_charge_amount:
      serviceCharge,

    total_amount:
      total,

    payment_status:
      d.payment_status ||
      'Pending',

    sale_date:
      dateValue(d.sale_date),

    created_by:
      d.created_by || ''

  };

  appendByHeader(
    'sales',
    record
  );

  return record;

}


function updateSale(d) {

  const row =
    findRow(
      'sales',
      'sales_id',
      d.sales_id
    );

  if (!row) {
    throw new Error(
      'Sale not found: ' +
      d.sales_id
    );
  }

  const quantity =
    num(d.quantity);

  const unitPrice =
    num(d.unit_price);

  const serviceRate =
    num(d.service_charge_rate);

  const commission =
    num(d.commission_amount);

  d.quantity =
    quantity;

  d.unit_price =
    unitPrice;

  d.service_charge_rate =
    serviceRate;

  d.commission_amount =
    commission;

  d.subtotal =
    quantity * unitPrice;

  d.service_charge_amount =
    d.subtotal *
    serviceRate /
    100;

  d.total_amount =
    d.subtotal -
    d.service_charge_amount -
    commission;

  setByHeader(
    'sales',
    row,
    d
  );

  return rowObject(
    'sales',
    row
  );

}


/* ==================================================
   COSTS
================================================== */

function addCost(d) {

  const id =
    nextId('costs');

  const quantity =
    num(d.quantity);

  const unitPrice =
    num(d.unit_price);

  const record = {

    cost_id:
      id,

    cost_code:
      d.cost_code ||
      code('CST', id),

    product_name:
      d.product_name || '',

    category:
      d.category || '',

    quantity:
      quantity,

    unit_price:
      unitPrice,

    total_price:
      quantity * unitPrice,

    cost_date:
      dateValue(d.cost_date),

    payment_status:
      d.payment_status ||
      'Unpaid',

    created_by:
      d.created_by || ''

  };

  appendByHeader(
    'costs',
    record
  );

  return record;

}


function updateCost(d) {

  const row =
    findRow(
      'costs',
      'cost_id',
      d.cost_id
    );

  if (!row) {
    throw new Error(
      'Cost not found: ' +
      d.cost_id
    );
  }

  const quantity =
    num(d.quantity);

  const unitPrice =
    num(d.unit_price);

  d.quantity =
    quantity;

  d.unit_price =
    unitPrice;

  d.total_price =
    quantity * unitPrice;

  setByHeader(
    'costs',
    row,
    d
  );

  return rowObject(
    'costs',
    row
  );

}


/* ==================================================
   CUSTOMERS
================================================== */

function addCustomer(d) {

  const id =
    nextId('customers');

  const record = {

    customer_id:
      id,

    customer_code:
      d.customer_code ||
      code('CUS', id),

    customer_name:
      d.customer_name || '',

    contact_number:
      d.contact_number || '',

    email:
      d.email || '',

    address:
      d.address || '',

    registered_date:
      dateValue(d.registered_date)

  };

  appendByHeader(
    'customers',
    record
  );

  return record;

}


function updateCustomer(d) {

  const row =
    findRow(
      'customers',
      'customer_id',
      d.customer_id
    );

  if (!row) {
    throw new Error(
      'Customer not found: ' +
      d.customer_id
    );
  }

  setByHeader(
    'customers',
    row,
    d
  );

  return rowObject(
    'customers',
    row
  );

}


/* ==================================================
   ORDERS
================================================== */

function addOrder(d) {

  const id =
    nextId('orders');

  const record = {

    order_id:
      id,

    order_code:
      d.order_code ||
      code('ORD', id),

    customer_id:
      d.customer_id || '',

    order_date:
      dateValue(d.order_date),

    quantity:
      num(d.quantity),

    order_status:
      d.order_status ||
      'Pending',

    estimated_delivery_date:
      d.estimated_delivery_date || '',

    delivery_address:
      d.delivery_address || '',

    created_by:
      d.created_by || ''

  };

  appendByHeader(
    'orders',
    record
  );

  return record;

}


function updateOrder(d) {

  const row =
    findRow(
      'orders',
      'order_id',
      d.order_id
    );

  if (!row) {
    throw new Error(
      'Order not found: ' +
      d.order_id
    );
  }

  d.quantity =
    num(d.quantity);

  setByHeader(
    'orders',
    row,
    d
  );

  return rowObject(
    'orders',
    row
  );

}


/* ==================================================
   PAYMENTS
================================================== */

function addPayment(d) {

  const id =
    nextId('payments');

  const record = {

    payment_id:
      id,

    payment_code:
      d.payment_code ||
      code('PAY', id),

    customer_id:
      d.customer_id || '',

    sales_id:
      d.sales_id || '',

    payment_date:
      dateValue(d.payment_date),

    payment_amount:
      num(d.payment_amount),

    payment_method:
      d.payment_method || '',

    transaction_ref:
      d.transaction_ref || '',

    payment_status:
      d.payment_status ||
      'Pending',

    created_by:
      d.created_by || ''

  };

  appendByHeader(
    'payments',
    record
  );

  return record;

}


function updatePayment(d) {

  const row =
    findRow(
      'payments',
      'payment_id',
      d.payment_id
    );

  if (!row) {
    throw new Error(
      'Payment not found: ' +
      d.payment_id
    );
  }

  d.payment_amount =
    num(d.payment_amount);

  setByHeader(
    'payments',
    row,
    d
  );

  return rowObject(
    'payments',
    row
  );

}


/* ==================================================
   RECEIPTS
================================================== */

function addReceipt(d) {

  const id =
    nextId('receipts');

  const record = {

    receipt_id:
      id,

    receipt_code:
      d.receipt_code ||
      code('RCP', id),

    payment_id:
      d.payment_id || '',

    sales_id:
      d.sales_id || '',

    customer_id:
      d.customer_id || '',

    amount_paid:
      num(d.amount_paid),

    total_sale_amount:
      num(d.total_sale_amount),

    total_paid:
      num(d.total_paid),

    balance_due:
      num(d.balance_due),

    payment_method:
      d.payment_method || '',

    payment_date:
      dateValue(d.payment_date),

    issued_by:
      d.issued_by || '',

    issued_at:
      dateValue(d.issued_at),

    status:
      d.status ||
      'Partial'

  };

  appendByHeader(
    'receipts',
    record
  );

  return record;

}


function updateReceipt(d) {

  const row =
    findRow(
      'receipts',
      'receipt_id',
      d.receipt_id
    );

  if (!row) {
    throw new Error(
      'Receipt not found: ' +
      d.receipt_id
    );
  }

  setByHeader(
    'receipts',
    row,
    d
  );

  return rowObject(
    'receipts',
    row
  );

}


/* ==================================================
   REPORTS
================================================== */

function addReport(d) {

  const id =
    nextId('reports');

  const totalGross =
    num(d.total_gross_sales);

  const totalService =
    num(d.total_service_charges);

  const totalCosts =
    num(d.total_costs);

  const commissions =
    num(d.total_commissions);

  const netProfit =
    d.net_profit !== undefined
      ? num(d.net_profit)
      : totalGross -
        totalService -
        totalCosts -
        commissions;

  const record = {

    report_id:
      id,

    report_code:
      d.report_code ||
      code('INC', id),

    report_month:
      d.report_month || '',

    report_year:
      d.report_year ||
      new Date().getFullYear(),

    total_gross_sales:
      totalGross,

    total_service_charges:
      totalService,

    total_costs:
      totalCosts,

    total_commissions:
      commissions,

    net_profit:
      netProfit,

    generated_by:
      d.generated_by || '',

    generated_at:
      dateValue(d.generated_at)

  };

  appendByHeader(
    'reports',
    record
  );

  return record;

}


/* ==================================================
   USERS
================================================== */

function addUser(d) {

  const username =
    String(d.username || '')
      .trim();

  if (!username) {
    throw new Error(
      'Username is required.'
    );
  }

  const existing =
    read('users').find(function(user) {

      return String(user.username)
        .trim()
        .toLowerCase() ===
        username.toLowerCase();

    });

  if (existing) {
    throw new Error(
      'Username already exists: ' +
      username
    );
  }

  const id =
    nextId('users');

  const record = {

    user_id:
      id,

    full_name:
      d.full_name || '',

    username:
      username,

    password:
      d.password || '',

    role:
      d.role || 'Staff',

    email:
      d.email || '',

    contact:
      d.contact || '',

    status:
      d.status || 'Active'

  };

  appendByHeader(
    'users',
    record
  );

  return record;

}


function updateUser(d) {

  const row =
    findRow(
      'users',
      'user_id',
      d.user_id
    );

  if (!row) {
    throw new Error(
      'User not found: ' +
      d.user_id
    );
  }

  setByHeader(
    'users',
    row,
    d
  );

  return rowObject(
    'users',
    row
  );

}


/* ==================================================
   LOGIN / USER ACTIVITY LOGS
================================================== */

function addLoginLog(d) {

  const id = nextId('login_logs');

  const record = {
    log_id: id,
    user_id: d.user_id !== undefined && d.user_id !== '' ? Number(d.user_id) : '',
    username: String(d.username || '').trim(),
    full_name: String(d.full_name || '').trim(),
    role: String(d.role || '').trim(),
    login_time: d.login_time || new Date(),
    action: String(d.action || 'LOGIN').toUpperCase(),
    status: String(d.status || 'SUCCESS').toUpperCase(),
    message: String(d.message || '').trim()
  };

  if (!record.username) {
    throw new Error('Username is required for login log.');
  }

  appendByHeader('login_logs', record);

  return record;

}


/* ==================================================
   PRODUCTS
================================================== */

function addProduct(d) {

  const id =
    nextId('products');

  const productCode =
    String(
      d.product_code ||
      ''
    ).trim();

  if (!productCode) {
    throw new Error(
      'Product code is required.'
    );
  }

  const duplicate =
    read('products').find(function(product) {

      return String(product.product_code)
        .trim()
        .toLowerCase() ===
        productCode.toLowerCase();

    });

  if (duplicate) {
    throw new Error(
      'Product code already exists: ' +
      productCode
    );
  }

  const record = {

    product_id:
      id,

    product_code:
      productCode,

    product_name:
      d.product_name ||
      productCode,

    unit_price:
      num(d.unit_price)

  };

  appendByHeader(
    'products',
    record
  );

  return record;

}


function updateProduct(d) {

  const row =
    findRow(
      'products',
      'product_id',
      d.product_id
    );

  if (!row) {
    throw new Error(
      'Product not found: ' +
      d.product_id
    );
  }

  if (d.unit_price !== undefined) {
    d.unit_price =
      num(d.unit_price);
  }

  setByHeader(
    'products',
    row,
    d
  );

  return rowObject(
    'products',
    row
  );

}


/* ==================================================
   DELETE
================================================== */

function deleteRecord(
  name,
  idField,
  id
) {

  const row =
    findRow(
      name,
      idField,
      id
    );

  if (!row) {

    throw new Error(
      name +
      ' record not found: ' +
      id
    );

  }

  sheet(name)
    .deleteRow(row);

  return true;

}


/* ==================================================
   REPLACE TABLE DATA
================================================== */

function replaceData(
  name,
  rows
) {

  if (!HEADERS[name]) {

    throw new Error(
      'Unsupported sheet: ' +
      name
    );

  }

  const sh =
    sheet(name);

  const headers =
    HEADERS[name];

  const lastRow =
    sh.getLastRow();

  if (lastRow > 1) {

    sh
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .clearContent();

  }

  if (
    Array.isArray(rows) &&
    rows.length
  ) {

    const values =
      rows.map(function(row) {

        return headers.map(
          function(header) {

            return row[header] !== undefined
              ? row[header]
              : '';

          }
        );

      });

    sh
      .getRange(
        2,
        1,
        values.length,
        headers.length
      )
      .setValues(values);

  }

  return true;

}


/* ==================================================
   DASHBOARD
================================================== */

function getDashboard() {

  const sales =
    read('sales');

  const costs =
    read('costs');

  let totalSales = 0;
  let totalCosts = 0;

  sales.forEach(function(sale) {

    totalSales +=
      num(sale.total_amount);

  });

  costs.forEach(function(cost) {

    totalCosts +=
      num(cost.total_price);

  });

  return {

    totalSales:
      totalSales,

    totalCosts:
      totalCosts,

    netProfit:
      totalSales -
      totalCosts,

    salesCount:
      sales.length,

    costCount:
      costs.length,

    customersCount:
      read('customers').length,

    ordersCount:
      read('orders').length,

    paymentsCount:
      read('payments').length,

    receiptsCount:
      read('receipts').length

  };

}


/* ==================================================
   SETUP SHEETS
================================================== */

function setupSheets() {

  const names =
    Object.keys(HEADERS);

  names.forEach(function(name) {

    sheet(name);

  });


  /* ---------------- USERS ---------------- */

  const users =
    read('users');

  if (!users.length) {

    appendByHeader(
      'users',
      {
        user_id: 1,
        full_name: 'System Admin',
        username: 'admin',
        password: 'admin123',
        role: 'Admin',
        email: 'admin@seemanfoods.lk',
        contact: '0771234567',
        status: 'Active'
      }
    );


    appendByHeader(
      'users',
      {
        user_id: 2,
        full_name: 'Kamal Perera',
        username: 'staff',
        password: 'staff123',
        role: 'Staff',
        email: 'kamal@seemanfoods.lk',
        contact: '0779876543',
        status: 'Active'
      }
    );

  }


  /* ---------------- PRODUCTS ---------------- */

  const products =
    read('products');

  const requiredProducts = [

    {
      product_code: 'SB - 001',
      product_name: 'SB - 001',
      unit_price: 0
    },

    {
      product_code: 'SB - 002',
      product_name: 'SB - 002',
      unit_price: 0
    },

    {
      product_code: 'SB - 003',
      product_name: 'SB - 003',
      unit_price: 0
    }

  ];


  requiredProducts.forEach(
    function(product) {

      const exists =
        products.some(
          function(existing) {

            return String(
              existing.product_code
            ).trim().toLowerCase() ===
            product.product_code
              .trim()
              .toLowerCase();

          }
        );

      if (!exists) {

        appendByHeader(
          'products',
          {
            product_id:
              nextId('products'),

            product_code:
              product.product_code,

            product_name:
              product.product_name,

            unit_price:
              product.unit_price

          }
        );

      }

    }
  );


  return 'Sheets ready';

}


/* ==================================================
   TEST FUNCTIONS
================================================== */

function testConnection() {

  return {

    success: true,

    message:
      'Google Apps Script connection is working.',

    spreadsheet:
      getSS().getName(),

    time:
      Utilities.formatDate(
        new Date(),
        TZ,
        'yyyy-MM-dd HH:mm:ss'
      )

  };

}


function testSetup() {

  const message =
    setupSheets();

  return {

    success: true,

    message:
      message,

    data:
      getAllData()

  };

}