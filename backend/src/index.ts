import express from 'express';
import cors from 'cors';
import { Product, Customer, Delivery, User, CompanyInfo } from './data/masterData';
import { Request } from 'express';

// カスタムリクエスト型を定義してreq.fileを認識させる
interface CustomRequest extends Request {
  file?: Express.Multer.File;
}

let customers: Customer[] = [];
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import csv from 'csv-parser';
import PDFDocument from 'pdfkit';

const app = express();
const port = 3002;

const dataDirectory = path.join(__dirname, 'data');
const invoicesFilePath = path.join(dataDirectory, 'invoices.json');
const deliveriesFilePath = path.join(dataDirectory, 'deliveries.json');
const customersFilePath = path.join(dataDirectory, 'customers.json');
const productsFilePath = path.join(dataDirectory, 'products.json');
const companyInfoFilePath = path.join(dataDirectory, 'companyInfo.json');

app.use(cors({
  origin: 'http://localhost:3000',
}));
app.use(express.json()); // JSONボディをパースするためのミドルウェア

const upload = multer({ dest: 'uploads/' });

// データ構造の定義
interface Invoice {
  id: string;
  voucherNumber: string;
  customerId: string;
  issueDate: string;
  items: { productId: string; quantity: number; unitPrice: number }[];
  totalAmount: number;
}



// データストレージ
let invoices: Invoice[] = [];
let deliveries: Delivery[] = [];
let products: Product[] = [];
let users: User[] = [];
let companyInfo: CompanyInfo = {
  name: '',
  postalCode: '',
  address: '',
  phone: '',
  fax: '',
  bankName: '',
  bankBranch: '',
  bankAccountType: '',
  bankAccountNumber: '',
  bankAccountHolder: '',
  contactPerson: '',
};
let currentVoucherNumber = 1;

// データをファイルから読み込む関数
const loadData = () => {
  try {
    if (fs.existsSync(invoicesFilePath)) {
      const invoicesData = fs.readFileSync(invoicesFilePath, 'utf-8');
      invoices = invoicesData ? JSON.parse(invoicesData) : [];
    } else {
      invoices = [];
    }
    if (fs.existsSync(deliveriesFilePath)) {
      const deliveriesData = fs.readFileSync(deliveriesFilePath, 'utf-8');
      deliveries = deliveriesData ? JSON.parse(deliveriesData) : [];
    } else {
      deliveries = [];
    }
    if (fs.existsSync(customersFilePath)) {
      const customersData = fs.readFileSync(customersFilePath, 'utf-8');
      customers = customersData ? JSON.parse(customersData) : [];
    } else {
      customers = [];
    }
    if (fs.existsSync(productsFilePath)) {
      const productsData = fs.readFileSync(productsFilePath, 'utf-8');
      products = productsData ? JSON.parse(productsData) : [];
    } else {
      products = [];
    }
    if (fs.existsSync(companyInfoFilePath)) {
      const companyInfoData = fs.readFileSync(companyInfoFilePath, 'utf-8');
      companyInfo = companyInfoData ? JSON.parse(companyInfoData) : companyInfo;
    } else {
      // If file doesn't exist, use default empty companyInfo
      companyInfo = {
        name: '',
        postalCode: '',
        address: '',
        phone: '',
        fax: '',
        bankName: '',
        bankBranch: '',
        bankAccountType: '',
        bankAccountNumber: '',
        bankAccountHolder: '',
        contactPerson: '',
      };
    }

    // Update currentVoucherNumber to avoid duplicates
    const maxInvoiceVoucher = Math.max(...invoices.map(i => parseInt(i.voucherNumber.substring(1))), 0);
    const maxDeliveryVoucher = Math.max(...deliveries.map(d => parseInt(d.voucherNumber.substring(1))), 0);
    currentVoucherNumber = Math.max(maxInvoiceVoucher, maxDeliveryVoucher) + 1;

  } catch (error) {
    console.error('Error loading data:', error instanceof Error ? error.message : String(error));
  }
};

// データをファイルに保存する関数
const saveData = () => {
  try {
    fs.writeFileSync(invoicesFilePath, JSON.stringify(invoices, null, 2));
    fs.writeFileSync(deliveriesFilePath, JSON.stringify(deliveries, null, 2));
    fs.writeFileSync(customersFilePath, JSON.stringify(customers, null, 2));
    fs.writeFileSync(productsFilePath, JSON.stringify(products, null, 2));
    fs.writeFileSync(companyInfoFilePath, JSON.stringify(companyInfo, null, 2));
  } catch (error) {
    console.error('Error saving data:', error instanceof Error ? error.message : String(error));
  }
};

// 伝票番号を生成する関数
const generateVoucherNumber = () => {
  const voucherNumber = `V${String(currentVoucherNumber).padStart(5, '0')}`;
  currentVoucherNumber++;
  return voucherNumber;
};


// Helper function to send data as Excel
const sendExcel = async (res: express.Response, data: any[], filename: string, columns?: any[]) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  if (data.length > 0) {
    // Use provided columns or generate from keys
    worksheet.columns = columns || Object.keys(data[0]).map(key => ({ header: key, key: key, width: 20 }));
    worksheet.addRows(data);
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);

  await workbook.xlsx.write(res);
  res.end();
};

// Helper function to send data as CSV (既存のCSV関数は残しておくが、Excelに切り替える)
const sendCsv = (res: express.Response, data: any[], filename: string) => {
  const json2csv = require('json-2-csv').json2csv; // 型定義の問題を回避
  json2csv(data, (err: any, csv: any) => {
    if (err) {
      res.status(500).send('Error generating CSV');
      return;
    }
    res.header('Content-Type', 'text/csv');
    res.attachment(filename);
    res.send(csv);
  });
};

// Helper function to filter products
const filterProducts = (query: any) => {
  const { productName, productName_matchType, unit, unit_matchType, postalCode, postalCode_matchType,
          shippingAddress, shippingAddress_matchType, customer, customer_matchType, notes, notes_matchType,
          shippingName, shippingName_matchType, minUnitPrice, maxUnitPrice } = query;

  let filteredProducts = products;

  if (productName) {
    const matchType = productName_matchType || 'partial';
    if (matchType === 'exact') {
      filteredProducts = filteredProducts.filter(p => p.name === productName);
    } else { // partial
      filteredProducts = filteredProducts.filter(p => p.name.includes(productName));
    }
  }
  if (unit) {
    const matchType = unit_matchType || 'partial';
    if (matchType === 'exact') {
      filteredProducts = filteredProducts.filter(p => p.unit === unit);
    } else { // partial
      filteredProducts = filteredProducts.filter(p => p.unit.includes(unit));
    }
  }
  if (postalCode) {
    const matchType = postalCode_matchType || 'partial';
    if (matchType === 'exact') {
      filteredProducts = filteredProducts.filter(p => p.postalCode === postalCode);
    } else { // partial
      filteredProducts = filteredProducts.filter(p => p.postalCode.includes(postalCode));
    }
  }
  if (shippingAddress) {
    const matchType = shippingAddress_matchType || 'partial';
    if (matchType === 'exact') {
      filteredProducts = filteredProducts.filter(p => p.shippingAddress === shippingAddress);
    } else { // partial
      filteredProducts = filteredProducts.filter(p => p.shippingAddress.includes(shippingAddress));
    }
  }
  if (customer) {
    const matchType = customer_matchType || 'partial';
    if (matchType === 'exact') {
      filteredProducts = filteredProducts.filter(p => p.customer === customer);
    } else { // partial
      filteredProducts = filteredProducts.filter(p => p.customer.includes(customer));
    }
  }
  if (notes) {
    const matchType = notes_matchType || 'partial';
    if (matchType === 'exact') {
      filteredProducts = filteredProducts.filter(p => p.notes === notes);
    } else { // partial
      filteredProducts = filteredProducts.filter(p => p.notes.includes(notes));
    }
  }
  if (shippingName) {
    const matchType = shippingName_matchType || 'partial';
    if (matchType === 'exact') {
      filteredProducts = filteredProducts.filter(p => p.shippingName === shippingName);
    } else { // partial
      filteredProducts = filteredProducts.filter(p => p.shippingName?.includes(shippingName));
    }
  }
  if (minUnitPrice) {
    filteredProducts = filteredProducts.filter(p => p.unitPrice >= parseFloat(minUnitPrice as string));
  }
  if (maxUnitPrice) {
    filteredProducts = filteredProducts.filter(p => p.unitPrice <= parseFloat(maxUnitPrice as string));
  }
  return filteredProducts;
};

// Helper function to filter deliveries
const filterDeliveries = (query: any) => {
  const { voucherNumber, startDate, endDate, customerId, productId, minQuantity, maxQuantity, minUnitPrice, maxUnitPrice,
          status, salesGroup, unit, orderId, notes, minAmount, maxAmount, invoiceStatus,
          shippingAddressName, shippingPostalCode, shippingAddressDetail } = query;

  let filteredDeliveries = deliveries;

  if (voucherNumber) {
    filteredDeliveries = filteredDeliveries.filter(d => d.voucherNumber && d.voucherNumber.includes(voucherNumber));
  }
  if (startDate) {
    filteredDeliveries = filteredDeliveries.filter(d => d.deliveryDate >= startDate);
  }
  if (endDate) {
    filteredDeliveries = filteredDeliveries.filter(d => d.deliveryDate <= endDate);
  }
  if (customerId) {
    filteredDeliveries = filteredDeliveries.filter(d => d.customerId === customerId);
  }
  if (productId) {
    // This is a simplified filter. A real implementation might need to check inside the items array.
    filteredDeliveries = filteredDeliveries.filter(d => d.items.some(item => item.productId === productId));
  }
  if (minQuantity) {
    filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity, 0) >= parseFloat(minQuantity as string));
  }
  if (maxQuantity) {
    filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity, 0) <= parseFloat(maxQuantity as string));
  }
  if (minUnitPrice) {
    // This is a simplified filter. A real implementation might need to check inside the items array.
    filteredDeliveries = filteredDeliveries.filter(d => d.items.some(item => item.unitPrice >= parseFloat(minUnitPrice as string)));
  }
  if (maxUnitPrice) {
    // This is a simplified filter. A real implementation might need to check inside the items array.
    filteredDeliveries = filteredDeliveries.filter(d => d.items.some(item => item.unitPrice <= parseFloat(maxUnitPrice as string)));
  }
  // Other filters like status, salesGroup, etc. would need to be adapted to the new data structure.
  if (status) {
    filteredDeliveries = filteredDeliveries.filter(d => d.status === status);
  }
  if (salesGroup) {
    filteredDeliveries = filteredDeliveries.filter(d => d.salesGroup && d.salesGroup.includes(salesGroup));
  }
  if (unit) {
    filteredDeliveries = filteredDeliveries.filter(d => d.items.some(item => item.unit === unit));
  }
  if (orderId) {
    filteredDeliveries = filteredDeliveries.filter(d => d.orderId && d.orderId.includes(orderId));
  }
  if (notes) {
    filteredDeliveries = filteredDeliveries.filter(d => d.notes && d.notes.includes(notes));
  }
  if (minAmount) {
    filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0) >= parseFloat(minAmount as string));
  }
  if (maxAmount) {
    filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0) <= parseFloat(maxAmount as string));
  }
  if (invoiceStatus) {
    filteredDeliveries = filteredDeliveries.filter(d => d.invoiceStatus === invoiceStatus);
  }
  if (shippingAddressName) {
    filteredDeliveries = filteredDeliveries.filter(d => d.shippingAddressName && d.shippingAddressName.includes(shippingAddressName));
  }
  if (shippingPostalCode) {
    filteredDeliveries = filteredDeliveries.filter(d => d.shippingPostalCode && d.shippingPostalCode.includes(shippingPostalCode));
  }
  if (shippingAddressDetail) {
    filteredDeliveries = filteredDeliveries.filter(d => d.shippingAddressDetail && d.shippingAddressDetail.includes(shippingAddressDetail));
  }

  return filteredDeliveries;
};

// Helper function to filter invoices
const filterInvoices = (query: any) => {
  const { voucherNumber, customerId, startDate, endDate } = query;
  let filteredInvoices = invoices;

  if (voucherNumber) {
    filteredInvoices = filteredInvoices.filter(i => i.voucherNumber.includes(voucherNumber));
  }
  if (customerId) {
    filteredInvoices = filteredInvoices.filter(i => i.customerId === customerId);
  }
  if (startDate) {
    filteredInvoices = filteredInvoices.filter(i => i.issueDate >= startDate);
  }
  if (endDate) {
    filteredInvoices = filteredInvoices.filter(i => i.issueDate <= endDate);
  }
  return filteredInvoices;
}

// Helper function to filter customers
const filterCustomers = (query: any) => {
  const { name, name_matchType, postalCode, postalCode_matchType, address, address_matchType,
          phone, phone_matchType, paymentTerms, paymentTerms_matchType, email, email_matchType,
          contactPerson, contactPerson_matchType, minClosingDay, maxClosingDay, invoiceDeliveryMethod } = query;

  let filteredCustomers = customers;

  if (name) {
    const matchType = name_matchType || 'partial';
    if (matchType === 'exact') {
      filteredCustomers = filteredCustomers.filter(c => c.name === name);
    } else { // partial
      filteredCustomers = filteredCustomers.filter(c => c.name.includes(name));
    }
  }
  if (postalCode) {
    const matchType = postalCode_matchType || 'partial';
    if (matchType === 'exact') {
      filteredCustomers = filteredCustomers.filter(c => c.postalCode === postalCode);
    } else { // partial
      filteredCustomers = filteredCustomers.filter(c => c.postalCode.includes(postalCode));
    }
  }
  if (address) {
    const matchType = address_matchType || 'partial';
    if (matchType === 'exact') {
      filteredCustomers = filteredCustomers.filter(c => c.address === address);
    } else { // partial
      filteredCustomers = filteredCustomers.filter(c => c.address.includes(address));
    }
  }
  if (phone) {
    const matchType = phone_matchType || 'partial';
    if (matchType === 'exact') {
      filteredCustomers = filteredCustomers.filter(c => c.phone === phone);
    } else { // partial
      filteredCustomers = filteredCustomers.filter(c => c.phone.includes(phone));
    }
  }
  if (paymentTerms) {
    const matchType = paymentTerms_matchType || 'partial';
    if (matchType === 'exact') {
      filteredCustomers = filteredCustomers.filter(c => c.paymentTerms === paymentTerms);
    } else { // partial
      filteredCustomers = filteredCustomers.filter(c => c.paymentTerms.includes(paymentTerms));
    }
  }
  if (email) {
    const matchType = email_matchType || 'partial';
    if (matchType === 'exact') {
      filteredCustomers = filteredCustomers.filter(c => c.email === email);
    } else { // partial
      filteredCustomers = filteredCustomers.filter(c => c.email.includes(email));
    }
  }
  if (contactPerson) {
    const matchType = contactPerson_matchType || 'partial';
    if (matchType === 'exact') {
      filteredCustomers = filteredCustomers.filter(c => c.contactPerson === contactPerson);
    } else { // partial
      filteredCustomers = filteredCustomers.filter(c => c.contactPerson.includes(contactPerson));
    }
  }
  if (minClosingDay) {
    filteredCustomers = filteredCustomers.filter(c => c.closingDay >= parseInt(minClosingDay as string));
  }
  if (maxClosingDay) {
    filteredCustomers = filteredCustomers.filter(c => c.closingDay <= parseInt(maxClosingDay as string));
  }
  if (invoiceDeliveryMethod) {
    const methods = (invoiceDeliveryMethod as string).split(',');
    filteredCustomers = filteredCustomers.filter(c => methods.includes(c.invoiceDeliveryMethod));
  }
  return filteredCustomers;
};

// Helper function to filter users
const filterUsers = (query: any) => {
  const { username, username_matchType, email, email_matchType, role } = query;

  let filteredUsers = users;

  if (username) {
    const matchType = username_matchType || 'partial';
    if (matchType === 'exact') {
      filteredUsers = filteredUsers.filter((u: User) => u.username === username);
    } else { // partial
      filteredUsers = filteredUsers.filter((u: User) => u.username.includes(username));
    }
  }
  if (email) {
    const matchType = email_matchType || 'partial';
    if (matchType === 'exact') {
      filteredUsers = filteredUsers.filter((u: User) => u.email === email);
    } else { // partial
      filteredUsers = filteredUsers.filter((u: User) => u.email.includes(email));
    }
  }
  if (role) {
    filteredUsers = filteredUsers.filter((u: User) => u.role === role);
  }
  return filteredUsers;
};


// CSV Export Endpoints
app.get('/api/export/products', async (req, res) => {
  console.log('Received export request for products.');
  const filteredProducts = filterProducts(req.query);
  const productsForExport = filteredProducts.map(p => ({
    id: p.id || '',
    name: p.name || '',
    unitPrice: p.unitPrice || 0,
    unit: p.unit || '',
    shippingAddress: p.shippingAddress || '',
    postalCode: p.postalCode || '',
    customer: p.customer || '',
    notes: p.notes || '',
    shippingName: p.shippingName || '',
  }));
  const productColumns = [
    { header: 'id', key: 'id', width: 10 },
    { header: 'name', key: 'name', width: 30 },
    { header: 'unitPrice', key: 'unitPrice', width: 15 },
    { header: 'unit', key: 'unit', width: 10 },
    { header: 'shippingAddress', key: 'shippingAddress', width: 50 },
    { header: 'postalCode', key: 'postalCode', width: 15 },
    { header: 'customer', key: 'customer', width: 20 },
    { header: 'notes', key: 'notes', width: 30 },
    { header: 'shippingName', key: 'shippingName', width: 30 },
  ];
  await sendExcel(res, productsForExport, 'products.xlsx', productColumns);
});

app.get('/api/export/customers', async (req, res) => {
  console.log('Received export request for customers.');
  const filteredCustomers = filterCustomers(req.query);
  const customerColumns = [
    { header: 'id', key: 'id', width: 10 },
    { header: 'name', key: 'name', width: 30 },
    { header: 'formalName', key: 'formalName', width: 30 },
    { header: 'postalCode', key: 'postalCode', width: 15 },
    { header: 'address', key: 'address', width: 50 },
    { header: 'phone', key: 'phone', width: 20 },
    { header: 'paymentTerms', key: 'paymentTerms', width: 20 },
    { header: 'email', key: 'email', width: 30 },
    { header: 'contactPerson', key: 'contactPerson', width: 20 },
    { header: 'closingDay', key: 'closingDay', width: 10 },
    { header: 'invoiceDeliveryMethod', key: 'invoiceDeliveryMethod', width: 20 },
  ];
  await sendExcel(res, filteredCustomers, 'customers.xlsx', customerColumns);
});

app.get('/api/export/deliveries', async (req, res) => {
  console.log('Received export request for deliveries.');
  const filteredDeliveries = filterDeliveries(req.query);
  const deliveriesWithNames = filteredDeliveries.map(delivery => {
    const customer = customers.find(c => c.id === delivery.customerId);
    const customerName = customer ? customer.name : '不明';
    const itemsWithNames = delivery.items.map(item => {
      const product = products.find(p => p.id === item.productId);
      const productName = product ? product.name : '不明';
      return { ...item, productName };
    });
    return { ...delivery, customerName, items: itemsWithNames };
  });
  const deliveriesForExport = deliveriesWithNames.flatMap(delivery => {
    return delivery.items.map(item => ({
      id: delivery.id,
      voucherNumber: delivery.voucherNumber,
      deliveryDate: delivery.deliveryDate,
      customerName: delivery.customerName,
      productName: item.productName,
      quantity: item.quantity,
      unitPrice: item.unitPrice,
      unit: item.unit,
      notes: delivery.notes,
      orderId: delivery.orderId,
      status: delivery.status,
      invoiceStatus: delivery.invoiceStatus,
      salesGroup: delivery.salesGroup,
      shippingAddressName: delivery.shippingAddressName,
      shippingPostalCode: delivery.shippingPostalCode,
      shippingAddressDetail: delivery.shippingAddressDetail,
    }));
  });
  await sendExcel(res, deliveriesForExport, 'deliveries.xlsx');
});

app.get('/api/export/users', async (req, res) => {
  console.log('Received export request for users.');
  const filteredUsers = filterUsers(req.query);
  await sendExcel(res, filteredUsers, 'users.xlsx');
});

app.get('/api/export/salesSummary', async (req, res) => {
  console.log('Received export request for salesSummary.');
  const filters = req.query;
  const filteredDeliveries = filterDeliveries(req.query);

  const salesByCustomer: { [key: string]: number } = {};
  filteredDeliveries.forEach(delivery => {
    const customerName = customers.find(c => c.id === delivery.customerId)?.name || '不明';
    const amount = delivery.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
    salesByCustomer[customerName] = (salesByCustomer[customerName] || 0) + amount;
  });
  const dataToExport = Object.keys(salesByCustomer).map(customerName => ({
    customerName,
    totalSales: salesByCustomer[customerName],
  }));
  await sendExcel(res, dataToExport, 'sales_summary.xlsx');
});

// JSON Filter Endpoints
app.get('/api/filter/products', (req, res) => {
  const filteredProducts = filterProducts(req.query);
  res.json(filteredProducts);
});

app.get('/api/filter/customers', (req, res) => {
  const filteredCustomers = filterCustomers(req.query);
  res.json(filteredCustomers);
});

app.get('/api/filter/deliveries', (req, res) => {
  const filteredDeliveries = filterDeliveries(req.query);
  res.json(filteredDeliveries);
});

app.get('/api/filter/users', (req, res) => {
  const filteredUsers = filterUsers(req.query);
  res.json(filteredUsers);
});

app.get('/api/filter/salesSummary', (req, res) => {
  const filteredDeliveries = filterDeliveries(req.query);

  const salesByCustomer: { [key: string]: number } = {};
  filteredDeliveries.forEach(delivery => {
    const customerName = customers.find(c => c.id === delivery.customerId)?.name || '不明';
    const amount = delivery.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
    salesByCustomer[customerName] = (salesByCustomer[customerName] || 0) + amount;
  });
  const dataToExport = Object.keys(salesByCustomer).map(customerName => ({
    customerName,
    totalSales: salesByCustomer[customerName],
  }));
  res.json({ summary: dataToExport, details: filteredDeliveries });
});

// New endpoints for creating invoices and deliveries
app.post('/api/invoices', (req, res) => {
  const { customerId, issueDate, items } = req.body;
  const totalAmount = items.reduce((sum: number, item: any) => sum + item.quantity * item.unitPrice, 0);
  const newInvoice: Invoice = {
    id: String(invoices.length + 1),
    voucherNumber: generateVoucherNumber(),
    customerId,
    issueDate,
    items,
    totalAmount,
  };
  invoices.push(newInvoice);
  saveData();
  res.status(201).json(newInvoice);
});

app.post('/api/deliveries', (req, res) => {
  const { customerName, deliveryDate, items } = req.body;

  const customer = customers.find(c => c.name === customerName);
  if (!customer) {
    return res.status(400).json({ message: 'Customer not found.' });
  }

  const processedItems = items.map((item: any) => {
    const product = products.find(p => p.name === item.productName);
    return {
      productId: product ? product.id : '',
      productName: item.productName, // 自由入力された商品名を保持
      quantity: item.quantity,
      unitPrice: item.unitPrice,
      unit: item.unit || (product ? product.unit : ''), // 商品マスタから単位を取得、なければ自由入力
      notes: item.notes || '',
    };
  });

  const newDelivery: Delivery = {
    id: String(deliveries.length + 1),
    voucherNumber: generateVoucherNumber(),
    customerId: customer.id,
    deliveryDate,
    items: processedItems,
    notes: '', // デフォルト値を追加
    status: '未発行', // デフォルト値を追加
    invoiceStatus: '未請求', // デフォルト値を追加
  };
  deliveries.push(newDelivery);
  saveData();
  res.status(201).json(newDelivery);
});

app.post('/api/customers', (req, res) => {
  const newCustomer = req.body;
  const maxIdNum = customers.reduce((max, customer) => {
    const idNum = parseInt(customer.id.replace('C', ''));
    return isNaN(idNum) ? max : Math.max(max, idNum);
  }, 0);
  newCustomer.id = 'C' + String(maxIdNum + 1).padStart(3, '0');
  customers.push(newCustomer);
  saveData();
  res.status(201).json(newCustomer);
});

app.delete('/api/customers/:id', (req, res) => {
  const { id } = req.params;
  const initialLength = customers.length;
  customers = customers.filter(customer => customer.id !== id);
  if (customers.length < initialLength) {
    saveData();
    res.status(200).json({ message: 'Customer deleted successfully.' });
  } else {
    res.status(404).json({ message: 'Customer not found.' });
  }
});

app.put('/api/customers/:id', (req, res) => {
  const { id } = req.params;
  const updatedCustomerData = req.body;
  const customerIndex = customers.findIndex(c => c.id === id);

  if (customerIndex > -1) {
    customers[customerIndex] = { ...customers[customerIndex], ...updatedCustomerData };
    saveData();
    res.status(200).json(customers[customerIndex]);
  } else {
    res.status(404).json({ message: 'Customer not found.' });
  }
});

app.get('/api/filter/invoices', (req, res) => {
  const filteredInvoices = filterInvoices(req.query);
  res.json(filteredInvoices);
});


app.get('/api/filter/invoices', (req, res) => {
  const filteredInvoices = filterInvoices(req.query);
  res.json(filteredInvoices);
});

app.post('/api/import/customers', upload.single('file'), async (req: CustomRequest, res) => {
  if (!req.file) {
    return res.status(400).json({ message: 'No file uploaded.' });
  }

  const filePath = req.file.path;
  const customersToImport: Customer[] = [];

  try {
    if (req.file.mimetype === 'text/csv') {
      // CSVファイルの処理
      await new Promise<void>((resolve, reject) => {
        fs.createReadStream(filePath)
          .pipe(csv())
          .on('data', (row) => {
            customersToImport.push({
              id: row.id,
              name: row.name,
              formalName: row.formalName,
              postalCode: row.postalCode,
              address: row.address,
              phone: row.phone,
              closingDay: parseInt(row.closingDay || '0'),
              paymentTerms: row.paymentTerms,
              email: row.email,
              contactPerson: row.contactPerson,
              invoiceDeliveryMethod: row.invoiceDeliveryMethod,
            });
          })
          .on('end', () => {
            resolve();
          })
          .on('error', (error) => {
            reject(error);
          });
      });
    } else if (req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      // Excelファイルの処理
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1);

      if (worksheet) {
        const headerMap: { [key: number]: string } = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
          headerMap[colNumber] = cell.text.trim();
        });

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header row

          const customerData: any = {};
          row.eachCell((cell, colNumber) => {
            const header = headerMap[colNumber];
            if (header) {
              customerData[header] = cell.text || '';
            }
          });

          customersToImport.push({
            id: customerData.id,
            name: customerData.name,
            formalName: customerData.formalName,
            postalCode: customerData.postalCode,
            address: customerData.address,
            phone: customerData.phone,
            closingDay: parseInt(customerData.closingDay || '0'),
            paymentTerms: customerData.paymentTerms,
            email: customerData.email,
            contactPerson: customerData.contactPerson,
            invoiceDeliveryMethod: customerData.invoiceDeliveryMethod,
          });
        });
      }
    } else {
      return res.status(400).json({ message: 'Unsupported file type.' });
    }

    // 既存の顧客データとマージ（重複は上書き）
    customersToImport.forEach(importedCustomer => {
      const existingIndex = customers.findIndex(c => c.id === importedCustomer.id);
      if (existingIndex > -1) {
        customers[existingIndex] = importedCustomer; // 上書き
      } else {
        customers.push(importedCustomer);
      }
    });

    saveData();
    res.status(200).json({ message: 'Customers imported successfully.', importedCount: customersToImport.length });
  } catch (error) {
    console.error('Error importing customers:', error);
    res.status(500).json({ message: 'Failed to import customers.', error: error instanceof Error ? error.message : String(error) });
  } finally {
    // アップロードされた一時ファイルを削除
    fs.unlink(filePath, (err) => {
      if (err) console.error('Error deleting temporary file:', err);
    });
  }
});

app.post('/api/import/products', upload.single('file'), async (req: CustomRequest, res) => {
  if (!req.file) {
    return res.status(400).json({ message: 'No file uploaded.' });
  }

  const filePath = req.file.path;
  const productsToImport: Product[] = [];

  try {
    if (req.file.mimetype === 'text/csv') {
      // CSVファイルの処理
      await new Promise<void>((resolve, reject) => {
        fs.createReadStream(filePath)
          .pipe(csv())
          .on('data', (row) => {
            productsToImport.push({
              id: row.id,
              name: row.name,
              unitPrice: parseFloat(row.unitPrice || '0'),
              unit: row.unit,
              shippingAddress: row.shippingAddress,
              postalCode: row.postalCode,
              customer: row.customer,
              notes: row.notes,
              shippingName: row.shippingName,
            });
          })
          .on('end', () => {
            resolve();
          })
          .on('error', (error) => {
            reject(error);
          });
      });
    } else if (req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      // Excelファイルの処理
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1);

      if (worksheet) {
        const headerMap: { [key: number]: string } = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
          headerMap[colNumber] = cell.text.trim();
        });

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header row

          const productData: any = {};
          row.eachCell((cell, colNumber) => {
            const header = headerMap[colNumber];
            if (header) {
              productData[header] = cell.text || '';
            }
          });

          productsToImport.push({
            id: productData.id,
            name: productData.name,
            unitPrice: parseFloat(productData.unitPrice || '0'),
            unit: productData.unit,
            shippingAddress: productData.shippingAddress,
            postalCode: productData.postalCode,
            customer: productData.customer,
            notes: productData.notes,
            shippingName: productData.shippingName,
          });
        });
      }
    } else {
      return res.status(400).json({ message: 'Unsupported file type.' });
    }

    // 既存の商品データとマージ（重複は上書き）
    productsToImport.forEach(importedProduct => {
      const existingIndex = products.findIndex(p => p.id === importedProduct.id);
      if (existingIndex > -1) {
        products[existingIndex] = importedProduct; // 上書き
      } else {
        products.push(importedProduct);
      }
    });

    saveData();
    res.status(200).json({ message: 'Products imported successfully.', importedCount: productsToImport.length });
  } catch (error) {
    console.error('Error importing products:', error);
    res.status(500).json({ message: 'Failed to import products.', error: error instanceof Error ? error.message : String(error) });
  } finally {
    // アップロードされた一時ファイルを削除
    fs.unlink(filePath, (err) => {
      if (err) console.error('Error deleting temporary file:', err);
    });
  }
});

app.post('/api/products', (req, res) => {
  const newProduct = req.body;
  const maxIdNum = products.reduce((max, product) => {
    const idNum = parseInt(product.id.replace('P', ''));
    return isNaN(idNum) ? max : Math.max(max, idNum);
  }, 0);
  newProduct.id = 'P' + String(maxIdNum + 1).padStart(3, '0');
  products.push(newProduct);
  saveData();
  res.status(201).json(newProduct);
});

app.put('/api/products/:id', (req, res) => {
  const { id } = req.params;
  const updatedProductData = req.body;
  const productIndex = products.findIndex(p => p.id === id);

  if (productIndex > -1) {
    products[productIndex] = { ...products[productIndex], ...updatedProductData };
    saveData();
    res.status(200).json(products[productIndex]);
  } else {
    res.status(404).json({ message: 'Product not found.' });
  }
});

app.put('/api/deliveries/:id', (req, res) => {
  const { id } = req.params;
  const updatedDeliveryData = req.body;
  const deliveryIndex = deliveries.findIndex(d => d.id === id);

  if (deliveryIndex > -1) {
    try {
      const currentDelivery = deliveries[deliveryIndex];

      // customerNameからcustomerIdに変換
      let updatedCustomerId = currentDelivery.customerId; // デフォルトは現在のID
      if (updatedDeliveryData.customerName) {
        const customer = customers.find(c => c.name === updatedDeliveryData.customerName);
        if (customer) {
          updatedCustomerId = customer.id;
        } else {
          return res.status(400).json({ message: `Customer not found: ${updatedDeliveryData.customerName}` });
        }
      } else if (updatedDeliveryData.customerId) {
        updatedCustomerId = updatedDeliveryData.customerId;
      }

      const updatedItems = updatedDeliveryData.items ? updatedDeliveryData.items.map((item: any) => {
        let productIdToUse = item.productId;
        let productNameToUse = item.productName;
        let unitToUse = item.unit;
        let unitPriceToUse = item.unitPrice;

        // If productName is provided and productId is not, treat as free-form
        if (productNameToUse && !productIdToUse) {
          const existingProduct = products.find(p => p.name === productNameToUse);
          if (existingProduct) {
            productIdToUse = existingProduct.id;
            unitToUse = unitToUse || existingProduct.unit;
            unitPriceToUse = unitPriceToUse || existingProduct.unitPrice;
          } else {
            productIdToUse = ''; // Indicate free-form product
          }
        } else if (productIdToUse) {
          const masterProduct = products.find(p => p.id === productIdToUse);
          if (masterProduct) {
            productNameToUse = masterProduct.name;
            unitToUse = unitToUse || masterProduct.unit;
            unitPriceToUse = unitPriceToUse || masterProduct.unitPrice;
          } else {
            // productIdが見つからない場合もエラーを返す
            return res.status(400).json({ message: `Product not found with ID: ${productIdToUse}` });
          }
        }

        return {
          ...item,
          productId: productIdToUse,
          productName: productNameToUse,
          unit: unitToUse,
          unitPrice: unitPriceToUse,
        };
      }) : currentDelivery.items;

      deliveries[deliveryIndex] = {
        ...currentDelivery,
        ...updatedDeliveryData,
        items: updatedItems,
        customerId: updatedCustomerId,
      };
      saveData();
      res.status(200).json(deliveries[deliveryIndex]);
    } catch (error) {
      console.error('Error updating delivery:', error);
      res.status(500).json({ message: 'Failed to update delivery.', error: error instanceof Error ? error.message : String(error) });
    }
  } else {
    res.status(404).json({ message: 'Delivery not found.' });
  }
});



app.get('/api/deliveries/:id/pdf', (req, res) => {
  const { id } = req.params;
  const deliveryIndex = deliveries.findIndex(d => d.id === id);

  if (deliveryIndex === -1) {
    return res.status(404).send('Delivery not found');
  }

  const delivery = deliveries[deliveryIndex];
  const customer = customers.find(c => c.id === delivery.customerId);

  if (!customer) {
    return res.status(404).send('Customer not found');
  }

  try {
    const doc = new PDFDocument({ size: 'A4', margin: 0 }); // 余白を0に設定

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename=delivery_${id}.pdf`);

    doc.pipe(res);

    const fontPath = path.join(__dirname, 'src/fonts/NotoSansJP-Regular.ttf');
    doc.font(fontPath);

    const drawDeliveryNote = (isCopy: boolean, startY: number) => {
      const marginX = 30;
      const contentWidth = 595 - 2 * marginX; // A4 width - margins

      // Border (Removed)

      // Title
      doc.fontSize(16).text(`納　品　書（${isCopy ? '控' : 'お客様用'}）`, marginX, startY + 20, { align: 'center', width: contentWidth });

      // Top Right Block (Company Info & Dates)
      let currentYRight = startY + 20;
      const rightBlockX = marginX + contentWidth - 250;
      const rightBlockWidth = 240;

      // Delivery Note Number
      doc.fontSize(8).text(`No：${delivery.voucherNumber}`, rightBlockX, currentYRight, {
          width: rightBlockWidth,
          align: 'right'
      });
      currentYRight += 15;

      // Issue Date with underline
      doc.fontSize(8).text(`発行日: ${new Date().toLocaleDateString()}`, rightBlockX, currentYRight, {
          width: rightBlockWidth,
          align: 'right',
          underline: true
      });
      currentYRight += 25; // Add more space after the date

      // Company Info
      doc.font(path.join(__dirname, 'src/fonts/NotoSansJP-Bold.ttf')).fontSize(10).text(companyInfo.name, rightBlockX, currentYRight, { width: rightBlockWidth, align: 'right' });
      doc.font(path.join(__dirname, 'src/fonts/NotoSansJP-Regular.ttf')); // Reset to regular font
      currentYRight += 12;

      doc.fontSize(8); // Set font size for calculation

      // Determine the actual starting X position for the address text
      const addressText = companyInfo.address;
      const addressCalculatedWidth = doc.widthOfString(addressText);
      let addressActualStartX = rightBlockX; // Default to left edge of the block if it wraps or is long

      if (addressCalculatedWidth <= rightBlockWidth) {
        // If the address fits in one line, calculate its actual start X for right alignment
        addressActualStartX = rightBlockX + rightBlockWidth - addressCalculatedWidth;
      }

      // Postal Code aligned with the start of the address
      doc.text(`〒${companyInfo.postalCode}`, addressActualStartX, currentYRight, { width: rightBlockWidth, align: 'left' });
      currentYRight += 12;

      // Address (right-aligned as before)
      doc.text(companyInfo.address, rightBlockX, currentYRight, { width: rightBlockWidth, align: 'right' });
      currentYRight += 12;
      doc.fontSize(8).text(`TEL：${companyInfo.phone}`, rightBlockX, currentYRight, { width: rightBlockWidth, align: 'right' });
      currentYRight += 12;
      doc.fontSize(8).text(`FAX：${companyInfo.fax}`, rightBlockX, currentYRight, { width: rightBlockWidth, align: 'right' });
      currentYRight += 12;
      doc.fontSize(8).text(`担当者：${companyInfo.contactPerson}`, rightBlockX, currentYRight, { width: rightBlockWidth, align: 'right' });
      currentYRight += 12;
      doc.fontSize(8).text(`${companyInfo.bankName} ${companyInfo.bankBranch} ${companyInfo.bankAccountType} ${companyInfo.bankAccountNumber} ${companyInfo.bankAccountHolder}`, rightBlockX, currentYRight, { width: rightBlockWidth, align: 'right' });

      // Top Left Block (Customer Info)
      let currentYLeft = startY + 70;
      const leftBlockX = marginX + 10;

      doc.fontSize(8);
      doc.text(`お客様コードNo：${customer.id}`, leftBlockX, currentYLeft);
      currentYLeft += 12;
      doc.text(`〒${customer.postalCode}`, leftBlockX, currentYLeft);
      currentYLeft += 12; // Add space here
      doc.text(`${customer.address}`, leftBlockX, currentYLeft);
      currentYLeft += 24; // Add 1 extra line break (1 * 12)
      doc.fontSize(11).text(`${customer.formalName || customer.name}　御中`, leftBlockX, currentYLeft);

      // Message (Adjust Y based on the lower of the two top blocks)
      const messageY = Math.max(currentYLeft, currentYRight) + 20; // Take the max Y from both blocks and add some padding
      doc.fontSize(8).text('下記の通り納品致しましたのでご査収ください。', marginX + 10, messageY, { align: 'right', width: contentWidth - 10 });

      // Items Table (Adjust Y based on messageY)
      const tableTop = messageY + 13; // Adjusted Y to reduce space
      const col1X = marginX + 10; // 品番・品名
      const col2X = col1X + 150; // 数量
      const col3X = col2X + 50;  // 単位
      const col4X = col3X + 50;  // 単価
      const col5X = col4X + 70;  // 金額
      const col6X = col5X + 70;  // 備考

      const col1Width = 140;
      const col2Width = 40;
      const col3Width = 40;
      const col4Width = 60;
      const col5Width = 60;
      const col6Width = 100;

      const rowHeight = 14; // Adjusted for more rows
      const headerY = tableTop;
      let currentY = headerY + rowHeight;

      // Table Headers
      doc.rect(col1X - 5, headerY, contentWidth - 10, rowHeight).fillAndStroke('#EEEEEE', '#000000');
      doc.fillColor('black').fontSize(8); // Adjusted font size
      doc.text('品番・品名', col1X, headerY + 3, { width: col1Width, align: 'center' });
      doc.text('数量', col2X, headerY + 3, { width: col2Width, align: 'center' });
      doc.text('単位', col3X, headerY + 3, { width: col3Width, align: 'center' });
      doc.text('単価', col4X, headerY + 3, { width: col4Width, align: 'center' });
      doc.text('金額', col5X, headerY + 3, { width: col5Width, align: 'center' });
      doc.text('備考', col6X, headerY + 3, { width: col6Width, align: 'center' });

      let totalAmount = 0;

      delivery.items.forEach(item => {
        const product = products.find(p => p.id === item.productId);
        const productName = product ? product.name : item.productName || '不明な商品';
        const amount = (item.quantity || 0) * (item.unitPrice || 0);
        totalAmount += amount;

        doc.rect(col1X - 5, currentY, contentWidth - 10, rowHeight).stroke();
        doc.fillColor('black').fontSize(8); // Adjusted font size
        doc.text(productName, col1X, currentY + 3, { width: col1Width, align: 'left' });
        doc.text(item.quantity?.toLocaleString() || '', col2X, currentY + 3, { width: col2Width, align: 'right' });
        doc.text(item.unit || '', col3X, currentY + 3, { width: col3Width, align: 'center' });
        doc.text(item.unitPrice?.toLocaleString() || '', col4X, currentY + 3, { width: col4Width, align: 'right' });
        doc.text(amount.toLocaleString(), col5X, currentY + 3, { width: col5Width, align: 'right' });
        doc.text(item.notes || '', col6X, currentY + 3, { width: col6Width, align: 'center' });
        currentY += rowHeight;
      });

      // Fill remaining rows if less than 12 items
      const minRows = 12;
      for (let i = delivery.items.length; i < minRows; i++) {
        doc.rect(col1X - 5, currentY, contentWidth - 10, rowHeight).stroke();
        currentY += rowHeight;
      }

      const tableBottomY = currentY; // This is the bottom of the last row drawn

      // Total row background
      doc.rect(col1X - 5, tableBottomY, contentWidth - 10, rowHeight).fillAndStroke('#EEEEEE', '#000000');
      doc.fillColor('black').fontSize(8);
      doc.text('合計', col4X, tableBottomY + 3, { width: col4Width, align: 'center' }); // '合計' in '単価' column
      doc.text(totalAmount.toLocaleString(), col5X, tableBottomY + 3, { width: col5Width, align: 'right' }); // Total amount in '金額' column

      // Draw vertical lines for the entire table, extending to the bottom of the total row
      doc.moveTo(col1X - 5, headerY)
         .lineTo(col1X - 5, tableBottomY + rowHeight) // Leftmost line
         .moveTo(col2X - 5, headerY)
         .lineTo(col2X - 5, tableBottomY + rowHeight) // Line between col1 and col2
         .moveTo(col3X - 5, headerY)
         .lineTo(col3X - 5, tableBottomY + rowHeight) // Line between col2 and col3
         .moveTo(col4X - 5, headerY)
         .lineTo(col4X - 5, tableBottomY + rowHeight) // Line between col3 and col4
         .moveTo(col5X - 5, headerY)
         .lineTo(col5X - 5, tableBottomY + rowHeight) // Line between col4 and col5
         .moveTo(col6X - 5, headerY)
         .lineTo(col6X - 5, tableBottomY + rowHeight) // Line between col5 and col6
         .moveTo(col1X - 5 + contentWidth - 10, headerY)
         .lineTo(col1X - 5 + contentWidth - 10, tableBottomY + rowHeight) // Rightmost line
         .stroke();

      // 消費税等は「請求書」で一括請求させて頂きます。
      doc.fontSize(8).text('消費税等は「請求書」で一括請求させて頂きます。', marginX + 10, tableBottomY + rowHeight + 2);
    };

    // Draw first delivery note (控)
    drawDeliveryNote(true, 0);

    // Draw second delivery note (お客様用)
    drawDeliveryNote(false, 420); // A4 height is 842, so roughly half + some spacing

    doc.end();

    // Update delivery status to '発行済み' after successful PDF generation
    deliveries[deliveryIndex].status = '発行済み';
    saveData();

  } catch (error) {
    console.error('Error generating PDF:', error);
    res.status(500).send('Error generating PDF');
  }
});

app.get('/api/deliveries/:id/invoice-pdf', (req, res) => {
  const { id } = req.params;
  const deliveryIndex = deliveries.findIndex(d => d.id === id);

  if (deliveryIndex === -1) {
    return res.status(404).send('Delivery not found');
  }

  const delivery = deliveries[deliveryIndex];
  const customer = customers.find(c => c.id === delivery.customerId);

  if (!customer) {
    return res.status(404).send('Customer not found');
  }

  try {
    const doc = new PDFDocument({ size: 'A4', margin: 20 });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename=invoice_${id}.pdf`);

    doc.pipe(res);

    const fontPath = path.join(__dirname, '../src/fonts/NotoSansJP-Regular.ttf');
    doc.font(fontPath);

    // Header
    const text = '請   求   書';
    const extraHorizontalSpacePerSide = 3; // 左右にそれぞれ3文字分の余白
    doc.fontSize(16); // フォントサイズを16ptに設定
    const textWidth = doc.widthOfString(text);
    const textHeight = doc.currentLineHeight();
    const charWidth = textWidth / text.length; // 1文字あたりの幅を概算

    // ボックスの寸法を計算
    const boxWidth = textWidth + (extraHorizontalSpacePerSide * 2 * charWidth);
    const boxHeight = textHeight; // 縦方向は文字の高さに合わせる

    // 請求書タイトルブロックと発行日のY座標を決定
    const commonTopY = doc.page.margins.top; // 上部の余白を最小限に

    // ボックスの位置を計算（中央に配置）
    const boxX = (doc.page.width - boxWidth) / 2;
    const boxY = commonTopY; // 共通のY座標を使用

    // 黒い四角形を描画
    doc.rect(boxX, boxY, boxWidth, boxHeight).fill('black');

    // 白い文字でテキストを四角形の中央に描画
    doc.fillColor('white').text(text, boxX, boxY, {
        width: boxWidth,
        align: 'center'
    });

    // 後続のテキストのために文字色を黒に戻す
    doc.fillColor('black');

    // 発行日を追加 (右上)
    const issueDateText = `発行日: ${new Date().toLocaleDateString()}`;
    const issueDateFontSize = 8.4; // 請求書タイトルの6割のサイズ
    doc.fontSize(issueDateFontSize);

    // 発行日のY座標も共通のY座標を使用
    const issueDateY = commonTopY;

    // 発行日を右寄せで配置
    const currentYBeforeIssueDate = doc.y; // 現在のdoc.yを保存
    doc.y = issueDateY;
    doc.text(issueDateText, { align: 'right' });

    // 請求書番号を追加 (発行日のすぐ下、同じ8.4pt)
    const invoiceNumberText = `請求書番号: ${delivery.voucherNumber}`;
    doc.text(invoiceNumberText, { align: 'right' }); // doc.yは自動的に進む

    doc.y = currentYBeforeIssueDate; // doc.yを元の位置に戻す

    // タイトルブロックの後のコンテンツの開始位置を調整
    doc.y = commonTopY + boxHeight + 10; // タイトルブロックのすぐ下に10ptのスペース

    // Info
    doc.fontSize(8); // 取引先情報のフォントサイズを8ptに設定

    const leftContentX = doc.page.margins.left + 20; // 左マージンからさらに20pt右にずらす
    const contentAreaWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

    // Capture the starting Y position for the customer and company info block
    const infoBlockStartY = doc.y;

    // --- Customer Information (Left Side) ---
    let customerInfoCurrentY = infoBlockStartY;

    // Customer Postal Code (10pt)
    doc.fontSize(10); // Set font size to 10pt
    doc.text(`〒${customer.postalCode}`, leftContentX, customerInfoCurrentY);
    customerInfoCurrentY += doc.currentLineHeight(); // Update Y based on 10pt font height

    // Customer Address (10pt)
    doc.text(`${customer.address}`, leftContentX, customerInfoCurrentY);
    customerInfoCurrentY += doc.currentLineHeight(); // Update Y based on 10pt font height
    customerInfoCurrentY += doc.currentLineHeight(); // Add an extra line break (equivalent to 10pt height)

    // Customer Formal Name (12pt)
    doc.fontSize(12); // Set font size to 12pt
    doc.text(`${customer.formalName || customer.name}　　御中`, leftContentX, customerInfoCurrentY);
    customerInfoCurrentY += doc.currentLineHeight(); // Update Y based on 12pt font height
    doc.fontSize(8); // Reset font size to 8pt for subsequent lines

    // --- Company Information (Right Side, aligned with customer info's top) ---
    let companyInfoCurrentY = infoBlockStartY;

    // Calculate the left X for the company info block to be right-aligned
    const companyBlockWidth = 200; // Example fixed width for company info block
    const companyBlockLeftX = doc.page.width - doc.page.margins.right - companyBlockWidth;

    // Calculate the actual starting X for right-aligned address to align other fields
    const addressTextWidth = doc.widthOfString(companyInfo.address);
    const addressActualStartX = companyBlockLeftX + companyBlockWidth - addressTextWidth;

    // Company Name (12pt)
    doc.fontSize(12); // フォントサイズを12ptに設定
    doc.text(companyInfo.name, addressActualStartX, companyInfoCurrentY, { width: companyBlockWidth, align: 'left' }); // 左寄せ
    companyInfoCurrentY += doc.currentLineHeight();
    doc.fontSize(8); // フォントサイズを8ptに戻す

    // Company Postal Code
    doc.text(`〒${companyInfo.postalCode}`, addressActualStartX, companyInfoCurrentY, { width: companyBlockWidth, align: 'left' }); // 左寄せ
    companyInfoCurrentY += doc.currentLineHeight();

    // Company Address
    doc.text(companyInfo.address, companyBlockLeftX, companyInfoCurrentY, { width: companyBlockWidth, align: 'right' }); // Keep this right-aligned
    companyInfoCurrentY += doc.currentLineHeight();

    // Company Phone
    doc.text(`TEL: ${companyInfo.phone}`, addressActualStartX, companyInfoCurrentY, { width: companyBlockWidth, align: 'left' }); // 左寄せ
    companyInfoCurrentY += doc.currentLineHeight();

    // Company FAX
    doc.text(`FAX: ${companyInfo.fax}`, addressActualStartX, companyInfoCurrentY, { width: companyBlockWidth, align: 'left' }); // 左寄せ
    companyInfoCurrentY += doc.currentLineHeight();

    // Company Bank Account
    doc.text(`${companyInfo.bankName} ${companyInfo.bankBranch} ${companyInfo.bankAccountType} ${companyInfo.bankAccountNumber} ${companyInfo.bankAccountHolder}`, companyBlockLeftX, companyInfoCurrentY, { width: companyBlockWidth, align: 'right' }); // 右寄せ
    companyInfoCurrentY += doc.currentLineHeight();

    let yForRegistrationNumber = companyInfoCurrentY; // Capture Y before potential registration number
    if (companyInfo.invoiceRegistrationNumber) {
      doc.text(`登録番号: ${companyInfo.invoiceRegistrationNumber}`, addressActualStartX, companyInfoCurrentY, { width: companyBlockWidth, align: 'left' }); // 左寄せ
      companyInfoCurrentY += doc.currentLineHeight();
    }

    // Print customer ID at yForRegistrationNumber
    doc.fontSize(8).text(`お客様コードNo：${customer.id}`, leftContentX, yForRegistrationNumber, { width: companyBlockWidth, align: 'left' });

    // doc.y を、取引先情報と自社情報のブロックの最下部に合わせる
    // どちらかY座標が大きい方に合わせる
    doc.y = Math.max(doc.y, companyInfoCurrentY);

    // Add 2 line breaks after the customer/company info block
    doc.moveDown(2);

    // Add messages
    const messageLineY = doc.y; // Capture the current Y for this line
    doc.fontSize(8).text('毎度ありがとうございます。下記の通り御請求申し上げます。', leftContentX, messageLineY, { align: 'left' });
    doc.text('振込手数料は貴社にてご負担願います。', doc.page.margins.left, messageLineY, { align: 'right', width: contentAreaWidth });

    // Tables below messages
    let currentTableY = doc.y; // Start directly below the messages

    const tableSpacing = 0; // Space between the two tables
    const colWidth1 = 30;
    const colWidth2 = 68.5256;
    const colWidth3 = 68.5256;
    const singleTableWidth = colWidth1 + colWidth2 + colWidth3;
    const row1Height = 12;
    const row2Height = 28;
    const tableHeight = row1Height + row2Height;
    const cornerRadius = 5; // Radius for rounded corners

    // Function to draw a 2x3 table with rounded outer corners
    const drawTable = (startX: number, startY: number, rowData: string[]) => {
        // Draw the outer rounded rectangle for the entire table
        doc.roundedRect(startX, startY, singleTableWidth, tableHeight, cornerRadius).stroke();

        // Draw internal horizontal lines
        doc.moveTo(startX, startY + row1Height)
           .lineTo(startX + singleTableWidth, startY + row1Height)
           .stroke();

        // Draw internal vertical lines
        const vLine1X = startX + colWidth1;
        const vLine2X = startX + colWidth1 + colWidth2;
        doc.moveTo(vLine1X, startY)
           .lineTo(vLine1X, startY + tableHeight)
           .stroke();
        doc.moveTo(vLine2X, startY)
           .lineTo(vLine2X, startY + tableHeight)
           .stroke();

        // Add content to the top row
        doc.fontSize(8); // Set font size for table content
        const textPadding = 5;
        const topRowTextY = startY + (row1Height - doc.currentLineHeight()) / 2 + 1;
        doc.text('税率', startX + textPadding, topRowTextY, { width: colWidth1 - (textPadding * 2), align: 'center' });
        doc.text('対象金額計', vLine1X + textPadding, topRowTextY, { width: colWidth2 - (textPadding * 2), align: 'center' });
        doc.text('消費税等', vLine2X + textPadding, topRowTextY, { width: colWidth3 - (textPadding * 2), align: 'center' });

        // Add content to the second row
        const bottomRowTextY = startY + row1Height + (row2Height - doc.currentLineHeight()) / 2 + 1;
        doc.text(rowData[0], startX + textPadding, bottomRowTextY, { width: colWidth1 - (textPadding * 2), align: 'center' });
        doc.text(rowData[1], vLine1X + textPadding, bottomRowTextY, { width: colWidth2 - (textPadding * 2), align: 'center' });
        doc.text(rowData[2], vLine2X + textPadding, bottomRowTextY, { width: colWidth3 - (textPadding * 2), align: 'center' });

        return startY + tableHeight; // Return the Y position after drawing the table
    };

    // Draw first table
    drawTable(doc.page.margins.left, currentTableY, ['10', '', '']);

    // Draw second table (next to the first one)
    drawTable(doc.page.margins.left + singleTableWidth + tableSpacing, currentTableY, ['8', '', '']);

    // Draw the new 2x1 table on the right
    const newTableReferenceText = '振込手数料は貴社にてご負担願います。';
    doc.fontSize(8); // Ensure correct font size for width calculation
    const newTableWidth = doc.widthOfString(newTableReferenceText);
    const newTableX = doc.page.width - doc.page.margins.right - newTableWidth;

    // Draw outer rectangle and horizontal divider
    doc.roundedRect(newTableX, currentTableY, newTableWidth, tableHeight, cornerRadius).stroke();
    doc.moveTo(newTableX, currentTableY + row1Height)
       .lineTo(newTableX + newTableWidth, currentTableY + row1Height)
       .stroke();

    // Fill top cell with black
    doc.rect(newTableX + 1, currentTableY + 1, newTableWidth - 2, row1Height - 2).fill('black');

    // Add content to the top row (white text)
    const newTableTopTextY = currentTableY + (row1Height - doc.currentLineHeight()) / 2 + 1;
    doc.fillColor('white').text('今回御請求額', newTableX, newTableTopTextY, { width: newTableWidth, align: 'center' });

    // Reset color and add content to the bottom row
    const newTableBottomTextY = currentTableY + row1Height + (row2Height - doc.currentLineHeight()) / 2 + 1;
    doc.fillColor('black').text('', newTableX, newTableBottomTextY, { width: newTableWidth, align: 'center' });

    // Detailed Items Table
    const detailTableY = doc.y + 15;
    const detailTableLeftX = doc.page.margins.left;
    const detailTableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const detailRowHeight = 20;
    doc.fontSize(8);

    // Define Columns
    const cols = [
        { x: detailTableLeftX, width: 48, header: '伝票日付', key: 'deliveryDate' },
        { x: detailTableLeftX + 48, width: 48, header: '伝票No.', key: 'voucherNumber' },
        { x: detailTableLeftX + 96, width: 164, header: '品番・品名', key: 'product' },
        { x: detailTableLeftX + 260, width: 40, header: '数量', key: 'quantity' },
        { x: detailTableLeftX + 300, width: 30, header: '単位', key: 'unit' },
        { x: detailTableLeftX + 330, width: 40, header: '単価', key: 'unitPrice' },
        { x: detailTableLeftX + 370, width: 30, header: '税区分', key: 'taxRate' },
        { x: detailTableLeftX + 400, width: 55, header: '税抜金額', key: 'amount' },
        { x: detailTableLeftX + 455, width: detailTableWidth - 455, header: '備考', key: 'notes' }
    ];

    // Draw Header
    let currentY = detailTableY;
    doc.rect(detailTableLeftX, currentY, detailTableWidth, detailRowHeight).fillAndStroke('#DDDDDD', '#000000');
    doc.fillColor('black');
    cols.forEach((col, index) => {
        doc.text(col.header, col.x + 2, currentY + 6, { width: col.width - 4, align: 'center' });
        if (index < cols.length - 1) { // Draw vertical line for all but the last column
            doc.moveTo(col.x + col.width, currentY)
               .lineTo(col.x + col.width, currentY + detailRowHeight)
               .stroke('#000000');
        }
    });
    currentY += detailRowHeight;

    // Draw Rows
    let totalAmount = 0;
    delivery.items.forEach(item => {
        const product = products.find(p => p.id === item.productId);
        const amount = (item.quantity || 0) * (item.unitPrice || 0);
        totalAmount += amount;

        const rowData = {
            deliveryDate: delivery.deliveryDate,
            voucherNumber: delivery.voucherNumber,
            product: product ? product.name : (item.productName || '不明な商品'),
            quantity: item.quantity?.toLocaleString() || '',
            unit: item.unit || '',
            unitPrice: item.unitPrice?.toLocaleString() || '',
            taxRate: '10%', // Placeholder
            amount: amount.toLocaleString(),
            notes: item.notes || ''
        };

        // Calculate row height based on product name length
        const productTextHeight = doc.heightOfString(rowData.product, { width: cols[2].width - 4 });
        const dynamicRowHeight = Math.max(detailRowHeight, productTextHeight + 8);

        doc.rect(detailTableLeftX, currentY, detailTableWidth, dynamicRowHeight).stroke('#000000');
        cols.forEach(col => {
            const key = col.key as keyof typeof rowData;
            doc.text(rowData[key], col.x + 2, currentY + 4, {
                width: col.width - 4,
                align: (key === 'deliveryDate' || key === 'voucherNumber' || key === 'quantity' || key === 'unitPrice' || key === 'amount') ? 'right' :
                       (key === 'taxRate' || key === 'unit') ? 'center' : 'left'
            });
        });

        // Draw vertical lines for the row
        cols.forEach(col => {
            doc.moveTo(col.x, currentY).lineTo(col.x, currentY + dynamicRowHeight).stroke('#000000');
        });
        doc.moveTo(detailTableLeftX + detailTableWidth, currentY).lineTo(detailTableLeftX + detailTableWidth, currentY + dynamicRowHeight).stroke('#000000');


        currentY += dynamicRowHeight;
    });

    // Fill remaining space with empty rows
    const bottomMargin = doc.page.margins.bottom;
    const availableHeightForEmptyRows = doc.page.height - bottomMargin - currentY;
    const emptyRowHeight = detailRowHeight; // Use the base row height for empty rows
    const numberOfEmptyRows = Math.floor(availableHeightForEmptyRows / emptyRowHeight);

    for (let i = 0; i < numberOfEmptyRows; i++) {
        doc.rect(detailTableLeftX, currentY, detailTableWidth, emptyRowHeight).stroke('#000000');
        // Draw vertical lines for empty row
        cols.forEach(col => {
            doc.moveTo(col.x, currentY).lineTo(col.x, currentY + emptyRowHeight).stroke('#000000');
        });
        doc.moveTo(detailTableLeftX + detailTableWidth, currentY).lineTo(detailTableLeftX + detailTableWidth, currentY + emptyRowHeight).stroke('#000000');
        currentY += emptyRowHeight;
    }

    

    doc.end();

    // Update delivery status to '請求済み' after successful PDF generation
    deliveries[deliveryIndex].invoiceStatus = '請求済み';
    saveData();

  } catch (error) {
    console.error('Error generating Invoice PDF:', error);
    res.status(500).send('Error generating Invoice PDF');
  }
});







// Reset Endpoints
app.delete('/api/reset/products', (req, res) => {
  products = [];
  try {
    fs.unlinkSync(productsFilePath);
  } catch (error) {
    console.error('Error deleting products.json:', error);
  }
  saveData();
  res.status(200).json({ message: 'Products data reset successfully.' });
});

app.delete('/api/reset/customers', (req, res) => {
  customers = [];
  try {
    fs.unlinkSync(customersFilePath);
  } catch (error) {
    console.error('Error deleting customers.json:', error);
  }
  saveData();
  res.status(200).json({ message: 'Customers data reset successfully.' });
});

app.delete('/api/reset/deliveries', (req, res) => {
  deliveries = [];
  try {
    fs.unlinkSync(deliveriesFilePath);
  } catch (error) {
    console.error('Error deleting deliveries.json:', error);
  }
  saveData();
  res.status(200).json({ message: 'Deliveries data reset successfully.' });
});

app.delete('/api/reset/invoices', (req, res) => {
  invoices = []; // 請求書にはモックデータがないため空にする
  try {
    fs.unlinkSync(invoicesFilePath);
  } catch (error) {
    console.error('Error deleting invoices.json:', error);
  }
  saveData();
  res.status(200).json({ message: 'Invoices data reset successfully.' });
});

app.delete('/api/reset/users', (req, res) => {
  // ユーザーデータは通常リセットしないが、必要であれば実装
  // users = [...initialUsers]; // initialUsers が定義されていれば
  res.status(200).json({ message: 'Users data reset is not typically allowed or implemented this way.' });
});

app.post('/api/import/deliveries', upload.single('file'), async (req: CustomRequest, res) => {
  if (!req.file) {
    return res.status(400).json({ message: 'No file uploaded.' });
  }

  const filePath = req.file.path;
  const deliveriesMap: { [id: string]: Delivery } = {};

  try {
    if (req.file.mimetype === 'text/csv') {
      // CSVファイルの処理は現状維持（必要であれば同様のロジックを実装）
      return res.status(501).json({ message: 'CSV import for deliveries is not fully implemented with aggregation.' });
    } else if (req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1);

      if (worksheet) {
        const headerMap: { [key: number]: string } = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
          headerMap[colNumber] = cell.text.trim();
        });

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header row

          const deliveryData: any = {};
          row.eachCell((cell, colNumber) => {
            const header = headerMap[colNumber];
            if (header) {
              deliveryData[header] = cell.text || '';
            }
          });

          const customer = customers.find(c => c.name === (deliveryData.customerName || '').trim());
          const product = products.find(p => p.name === (deliveryData.productName || '').trim());

          const newItem = {
            productId: product ? product.id : '',
            quantity: parseInt(deliveryData.quantity || '0'),
            unitPrice: parseFloat(deliveryData.unitPrice || '0'),
            unit: deliveryData.unit || (product ? product.unit : ''),
          };

          const deliveryId = deliveryData.id || String(Math.max(...deliveries.map(d => parseInt(d.id || '0')), 0) + 1); // IDがない場合は自動生成

          if (deliveriesMap[deliveryId]) {
            deliveriesMap[deliveryId].items.push(newItem);
          } else {
            deliveriesMap[deliveryId] = {
              id: deliveryId,
              voucherNumber: deliveryData.voucherNumber,
              deliveryDate: deliveryData.deliveryDate,
              customerId: customer ? customer.id : '',
              items: [newItem],
              notes: deliveryData.notes,
              orderId: deliveryData.orderId,
              status: deliveryData.status as any,
              invoiceStatus: deliveryData.invoiceStatus as any,
              salesGroup: deliveryData.salesGroup,
              shippingAddressName: deliveryData.shippingAddressName,
              shippingPostalCode: deliveryData.shippingPostalCode,
              shippingAddressDetail: deliveryData.shippingAddressDetail,
            };
          }
        });
      }
    } else {
      return res.status(400).json({ message: 'Unsupported file type.' });
    }

    const deliveriesToImport = Object.values(deliveriesMap);

    deliveriesToImport.forEach(importedDelivery => {
      const existingIndex = deliveries.findIndex(d => d.id === importedDelivery.id);
      if (existingIndex > -1) {
        deliveries[existingIndex] = importedDelivery; // 上書き
      } else {
        deliveries.push(importedDelivery);
      }
    });

    saveData();
    res.status(200).json({ message: 'Deliveries imported successfully.', importedCount: deliveriesToImport.length });
  } catch (error) {
    console.error('Error importing deliveries:', error);
    res.status(500).json({ message: 'Failed to import deliveries.', error: error instanceof Error ? error.message : String(error) });
  } finally {
    fs.unlink(filePath, (err) => {
      if (err) console.error('Error deleting temporary file:', err);
    });
  }
});

// Company Info Endpoints
app.get('/api/company-info', (req, res) => {
  res.json(companyInfo);
});

app.post('/api/company-info', (req, res) => {
  companyInfo = { ...companyInfo, ...req.body };
  saveData();
  res.status(200).json(companyInfo);
});

app.listen(port, () => {
  loadData();
  console.log(`Backend server listening at http://localhost:${port}`);
});
