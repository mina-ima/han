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
    const doc = new PDFDocument({ size: 'A4', margin: 50 });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename=invoice_${id}.pdf`);

    doc.pipe(res);

    const fontPath = path.join(__dirname, '../src/fonts/NotoSansJP-Regular.ttf');
    doc.font(fontPath);

    // Header
    doc.fontSize(20).text('請求書', { align: 'center' });
    doc.moveDown();

    // Info
    doc.fontSize(12).text(`請求書番号: ${delivery.voucherNumber}`); // 納品番号を流用
    doc.text(`発行日: ${new Date().toLocaleDateString()}`);
    doc.text(`納品日: ${delivery.deliveryDate}`);
    doc.moveDown();

    // Customer Info
    doc.text(`顧客名: ${customer.name}`);
    doc.text(`住所: ${customer.address}`);
    doc.text(`電話番号: ${customer.phone}`);
    doc.moveDown();

    // Items Table
    const tableTop = doc.y;
    const itemX = 50;
    const quantityX = 250;
    const unitPriceX = 350;
    const amountX = 450;

    doc.fontSize(10).text('商品名', itemX, tableTop);
    doc.text('数量', quantityX, tableTop);
    doc.text('単価', unitPriceX, tableTop);
    doc.text('金額', amountX, tableTop);

    let y = tableTop + 25;
    let totalAmount = 0;

    delivery.items.forEach(item => {
      const product = products.find(p => p.id === item.productId);
      const productName = product ? product.name : '不明な商品';
      const amount = item.quantity * item.unitPrice;
      totalAmount += amount;

      doc.text(productName, itemX, y);
      doc.text(item.quantity.toString(), quantityX, y);
      doc.text(item.unitPrice.toString(), unitPriceX, y);
      doc.text(amount.toString(), amountX, y);
      y += 25;
    });

    // Total
    doc.moveDown();
    doc.fontSize(12).text(`合計金額: ${totalAmount.toLocaleString()} 円`, { align: 'right' });

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
