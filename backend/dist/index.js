"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
let customers = [];
const exceljs_1 = __importDefault(require("exceljs"));
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const multer_1 = __importDefault(require("multer"));
const csv_parser_1 = __importDefault(require("csv-parser"));
const app = (0, express_1.default)();
const port = 3002;
const dataDirectory = path_1.default.join(__dirname, 'data');
const invoicesFilePath = path_1.default.join(dataDirectory, 'invoices.json');
const deliveriesFilePath = path_1.default.join(dataDirectory, 'deliveries.json');
const customersFilePath = path_1.default.join(dataDirectory, 'customers.json');
const productsFilePath = path_1.default.join(dataDirectory, 'products.json');
app.use((0, cors_1.default)({
    origin: 'http://localhost:3000',
}));
app.use(express_1.default.json()); // JSONボディをパースするためのミドルウェア
const upload = (0, multer_1.default)({ dest: 'uploads/' });
// データストレージ
let invoices = [];
let deliveries = [];
let products = [];
let users = [];
let currentVoucherNumber = 1;
// データをファイルから読み込む関数
const loadData = () => {
    try {
        if (fs_1.default.existsSync(invoicesFilePath)) {
            const invoicesData = fs_1.default.readFileSync(invoicesFilePath, 'utf-8');
            invoices = invoicesData ? JSON.parse(invoicesData) : [];
        }
        else {
            invoices = [];
        }
        if (fs_1.default.existsSync(deliveriesFilePath)) {
            const deliveriesData = fs_1.default.readFileSync(deliveriesFilePath, 'utf-8');
            deliveries = deliveriesData ? JSON.parse(deliveriesData) : [];
        }
        else {
            deliveries = [];
        }
        if (fs_1.default.existsSync(customersFilePath)) {
            const customersData = fs_1.default.readFileSync(customersFilePath, 'utf-8');
            customers = customersData ? JSON.parse(customersData) : [];
        }
        else {
            customers = [];
        }
        if (fs_1.default.existsSync(productsFilePath)) {
            const productsData = fs_1.default.readFileSync(productsFilePath, 'utf-8');
            products = productsData ? JSON.parse(productsData) : [];
        }
        else {
            products = [];
        }
        // Update currentVoucherNumber to avoid duplicates
        const maxInvoiceVoucher = Math.max(...invoices.map(i => parseInt(i.voucherNumber.substring(1))), 0);
        const maxDeliveryVoucher = Math.max(...deliveries.map(d => parseInt(d.voucherNumber.substring(1))), 0);
        currentVoucherNumber = Math.max(maxInvoiceVoucher, maxDeliveryVoucher) + 1;
    }
    catch (error) {
        console.error('Error loading data:', error instanceof Error ? error.message : String(error));
    }
};
// データをファイルに保存する関数
const saveData = () => {
    try {
        fs_1.default.writeFileSync(invoicesFilePath, JSON.stringify(invoices, null, 2));
        fs_1.default.writeFileSync(deliveriesFilePath, JSON.stringify(deliveries, null, 2));
        fs_1.default.writeFileSync(customersFilePath, JSON.stringify(customers, null, 2));
        fs_1.default.writeFileSync(productsFilePath, JSON.stringify(products, null, 2));
    }
    catch (error) {
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
const sendExcel = (res, data, filename, columns) => __awaiter(void 0, void 0, void 0, function* () {
    const workbook = new exceljs_1.default.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    if (data.length > 0) {
        // Use provided columns or generate from keys
        worksheet.columns = columns || Object.keys(data[0]).map(key => ({ header: key, key: key, width: 20 }));
        worksheet.addRows(data);
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
    yield workbook.xlsx.write(res);
    res.end();
});
// Helper function to send data as CSV (既存のCSV関数は残しておくが、Excelに切り替える)
const sendCsv = (res, data, filename) => {
    const json2csv = require('json-2-csv').json2csv; // 型定義の問題を回避
    json2csv(data, (err, csv) => {
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
const filterProducts = (query) => {
    const { productName, productName_matchType, unit, unit_matchType, postalCode, postalCode_matchType, shippingAddress, shippingAddress_matchType, customer, customer_matchType, notes, notes_matchType, shippingName, shippingName_matchType, minUnitPrice, maxUnitPrice } = query;
    let filteredProducts = products;
    if (productName) {
        const matchType = productName_matchType || 'partial';
        if (matchType === 'exact') {
            filteredProducts = filteredProducts.filter(p => p.name === productName);
        }
        else { // partial
            filteredProducts = filteredProducts.filter(p => p.name.includes(productName));
        }
    }
    if (unit) {
        const matchType = unit_matchType || 'partial';
        if (matchType === 'exact') {
            filteredProducts = filteredProducts.filter(p => p.unit === unit);
        }
        else { // partial
            filteredProducts = filteredProducts.filter(p => p.unit.includes(unit));
        }
    }
    if (postalCode) {
        const matchType = postalCode_matchType || 'partial';
        if (matchType === 'exact') {
            filteredProducts = filteredProducts.filter(p => p.postalCode === postalCode);
        }
        else { // partial
            filteredProducts = filteredProducts.filter(p => p.postalCode.includes(postalCode));
        }
    }
    if (shippingAddress) {
        const matchType = shippingAddress_matchType || 'partial';
        if (matchType === 'exact') {
            filteredProducts = filteredProducts.filter(p => p.shippingAddress === shippingAddress);
        }
        else { // partial
            filteredProducts = filteredProducts.filter(p => p.shippingAddress.includes(shippingAddress));
        }
    }
    if (customer) {
        const matchType = customer_matchType || 'partial';
        if (matchType === 'exact') {
            filteredProducts = filteredProducts.filter(p => p.customer === customer);
        }
        else { // partial
            filteredProducts = filteredProducts.filter(p => p.customer.includes(customer));
        }
    }
    if (notes) {
        const matchType = notes_matchType || 'partial';
        if (matchType === 'exact') {
            filteredProducts = filteredProducts.filter(p => p.notes === notes);
        }
        else { // partial
            filteredProducts = filteredProducts.filter(p => p.notes.includes(notes));
        }
    }
    if (shippingName) {
        const matchType = shippingName_matchType || 'partial';
        if (matchType === 'exact') {
            filteredProducts = filteredProducts.filter(p => p.shippingName === shippingName);
        }
        else { // partial
            filteredProducts = filteredProducts.filter(p => { var _a; return (_a = p.shippingName) === null || _a === void 0 ? void 0 : _a.includes(shippingName); });
        }
    }
    if (minUnitPrice) {
        filteredProducts = filteredProducts.filter(p => p.unitPrice >= parseFloat(minUnitPrice));
    }
    if (maxUnitPrice) {
        filteredProducts = filteredProducts.filter(p => p.unitPrice <= parseFloat(maxUnitPrice));
    }
    return filteredProducts;
};
// Helper function to filter deliveries
const filterDeliveries = (query) => {
    const { voucherNumber, startDate, endDate, customerId, productId, minQuantity, maxQuantity, minUnitPrice, maxUnitPrice, status, salesGroup, unit, orderId, notes, minAmount, maxAmount, invoiceStatus, shippingAddressName, shippingPostalCode, shippingAddressDetail } = query;
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
        filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity, 0) >= parseFloat(minQuantity));
    }
    if (maxQuantity) {
        filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity, 0) <= parseFloat(maxQuantity));
    }
    if (minUnitPrice) {
        // This is a simplified filter. A real implementation might need to check inside the items array.
        filteredDeliveries = filteredDeliveries.filter(d => d.items.some(item => item.unitPrice >= parseFloat(minUnitPrice)));
    }
    if (maxUnitPrice) {
        // This is a simplified filter. A real implementation might need to check inside the items array.
        filteredDeliveries = filteredDeliveries.filter(d => d.items.some(item => item.unitPrice <= parseFloat(maxUnitPrice)));
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
        filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0) >= parseFloat(minAmount));
    }
    if (maxAmount) {
        filteredDeliveries = filteredDeliveries.filter(d => d.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0) <= parseFloat(maxAmount));
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
const filterInvoices = (query) => {
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
};
// Helper function to filter customers
const filterCustomers = (query) => {
    const { name, name_matchType, postalCode, postalCode_matchType, address, address_matchType, phone, phone_matchType, paymentTerms, paymentTerms_matchType, email, email_matchType, contactPerson, contactPerson_matchType, minClosingDay, maxClosingDay, invoiceDeliveryMethod } = query;
    let filteredCustomers = customers;
    if (name) {
        const matchType = name_matchType || 'partial';
        if (matchType === 'exact') {
            filteredCustomers = filteredCustomers.filter(c => c.name === name);
        }
        else { // partial
            filteredCustomers = filteredCustomers.filter(c => c.name.includes(name));
        }
    }
    if (postalCode) {
        const matchType = postalCode_matchType || 'partial';
        if (matchType === 'exact') {
            filteredCustomers = filteredCustomers.filter(c => c.postalCode === postalCode);
        }
        else { // partial
            filteredCustomers = filteredCustomers.filter(c => c.postalCode.includes(postalCode));
        }
    }
    if (address) {
        const matchType = address_matchType || 'partial';
        if (matchType === 'exact') {
            filteredCustomers = filteredCustomers.filter(c => c.address === address);
        }
        else { // partial
            filteredCustomers = filteredCustomers.filter(c => c.address.includes(address));
        }
    }
    if (phone) {
        const matchType = phone_matchType || 'partial';
        if (matchType === 'exact') {
            filteredCustomers = filteredCustomers.filter(c => c.phone === phone);
        }
        else { // partial
            filteredCustomers = filteredCustomers.filter(c => c.phone.includes(phone));
        }
    }
    if (paymentTerms) {
        const matchType = paymentTerms_matchType || 'partial';
        if (matchType === 'exact') {
            filteredCustomers = filteredCustomers.filter(c => c.paymentTerms === paymentTerms);
        }
        else { // partial
            filteredCustomers = filteredCustomers.filter(c => c.paymentTerms.includes(paymentTerms));
        }
    }
    if (email) {
        const matchType = email_matchType || 'partial';
        if (matchType === 'exact') {
            filteredCustomers = filteredCustomers.filter(c => c.email === email);
        }
        else { // partial
            filteredCustomers = filteredCustomers.filter(c => c.email.includes(email));
        }
    }
    if (contactPerson) {
        const matchType = contactPerson_matchType || 'partial';
        if (matchType === 'exact') {
            filteredCustomers = filteredCustomers.filter(c => c.contactPerson === contactPerson);
        }
        else { // partial
            filteredCustomers = filteredCustomers.filter(c => c.contactPerson.includes(contactPerson));
        }
    }
    if (minClosingDay) {
        filteredCustomers = filteredCustomers.filter(c => c.closingDay >= parseInt(minClosingDay));
    }
    if (maxClosingDay) {
        filteredCustomers = filteredCustomers.filter(c => c.closingDay <= parseInt(maxClosingDay));
    }
    if (invoiceDeliveryMethod) {
        const methods = invoiceDeliveryMethod.split(',');
        filteredCustomers = filteredCustomers.filter(c => methods.includes(c.invoiceDeliveryMethod));
    }
    return filteredCustomers;
};
// Helper function to filter users
const filterUsers = (query) => {
    const { username, username_matchType, email, email_matchType, role } = query;
    let filteredUsers = users;
    if (username) {
        const matchType = username_matchType || 'partial';
        if (matchType === 'exact') {
            filteredUsers = filteredUsers.filter((u) => u.username === username);
        }
        else { // partial
            filteredUsers = filteredUsers.filter((u) => u.username.includes(username));
        }
    }
    if (email) {
        const matchType = email_matchType || 'partial';
        if (matchType === 'exact') {
            filteredUsers = filteredUsers.filter((u) => u.email === email);
        }
        else { // partial
            filteredUsers = filteredUsers.filter((u) => u.email.includes(email));
        }
    }
    if (role) {
        filteredUsers = filteredUsers.filter((u) => u.role === role);
    }
    return filteredUsers;
};
// CSV Export Endpoints
app.get('/api/export/products', (req, res) => __awaiter(void 0, void 0, void 0, function* () {
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
    yield sendExcel(res, productsForExport, 'products.xlsx', productColumns);
}));
app.get('/api/export/customers', (req, res) => __awaiter(void 0, void 0, void 0, function* () {
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
    yield sendExcel(res, filteredCustomers, 'customers.xlsx', customerColumns);
}));
app.get('/api/export/deliveries', (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    console.log('Received export request for deliveries.');
    const filteredDeliveries = filterDeliveries(req.query);
    const deliveriesWithNames = filteredDeliveries.map(delivery => {
        const customer = customers.find(c => c.id === delivery.customerId);
        const customerName = customer ? customer.name : '不明';
        const itemsWithNames = delivery.items.map(item => {
            const product = products.find(p => p.id === item.productId);
            const productName = product ? product.name : '不明';
            return Object.assign(Object.assign({}, item), { productName });
        });
        return Object.assign(Object.assign({}, delivery), { customerName, items: itemsWithNames });
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
    yield sendExcel(res, deliveriesForExport, 'deliveries.xlsx');
}));
app.get('/api/export/users', (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    console.log('Received export request for users.');
    const filteredUsers = filterUsers(req.query);
    yield sendExcel(res, filteredUsers, 'users.xlsx');
}));
app.get('/api/export/salesSummary', (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    console.log('Received export request for salesSummary.');
    const filters = req.query;
    const filteredDeliveries = filterDeliveries(req.query);
    const salesByCustomer = {};
    filteredDeliveries.forEach(delivery => {
        var _a;
        const customerName = ((_a = customers.find(c => c.id === delivery.customerId)) === null || _a === void 0 ? void 0 : _a.name) || '不明';
        const amount = delivery.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
        salesByCustomer[customerName] = (salesByCustomer[customerName] || 0) + amount;
    });
    const dataToExport = Object.keys(salesByCustomer).map(customerName => ({
        customerName,
        totalSales: salesByCustomer[customerName],
    }));
    yield sendExcel(res, dataToExport, 'sales_summary.xlsx');
}));
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
    const salesByCustomer = {};
    filteredDeliveries.forEach(delivery => {
        var _a;
        const customerName = ((_a = customers.find(c => c.id === delivery.customerId)) === null || _a === void 0 ? void 0 : _a.name) || '不明';
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
    const totalAmount = items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
    const newInvoice = {
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
    const processedItems = items.map((item) => {
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
    const newDelivery = {
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
    }
    else {
        res.status(404).json({ message: 'Customer not found.' });
    }
});
app.put('/api/customers/:id', (req, res) => {
    const { id } = req.params;
    const updatedCustomerData = req.body;
    const customerIndex = customers.findIndex(c => c.id === id);
    if (customerIndex > -1) {
        customers[customerIndex] = Object.assign(Object.assign({}, customers[customerIndex]), updatedCustomerData);
        saveData();
        res.status(200).json(customers[customerIndex]);
    }
    else {
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
app.post('/api/import/customers', upload.single('file'), (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    if (!req.file) {
        return res.status(400).json({ message: 'No file uploaded.' });
    }
    const filePath = req.file.path;
    const customersToImport = [];
    try {
        if (req.file.mimetype === 'text/csv') {
            // CSVファイルの処理
            yield new Promise((resolve, reject) => {
                fs_1.default.createReadStream(filePath)
                    .pipe((0, csv_parser_1.default)())
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
        }
        else if (req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            // Excelファイルの処理
            const workbook = new exceljs_1.default.Workbook();
            yield workbook.xlsx.readFile(filePath);
            const worksheet = workbook.getWorksheet(1);
            if (worksheet) {
                const headerMap = {};
                worksheet.getRow(1).eachCell((cell, colNumber) => {
                    headerMap[colNumber] = cell.text.trim();
                });
                worksheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1)
                        return; // Skip header row
                    const customerData = {};
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
        }
        else {
            return res.status(400).json({ message: 'Unsupported file type.' });
        }
        // 既存の顧客データとマージ（重複は上書き）
        customersToImport.forEach(importedCustomer => {
            const existingIndex = customers.findIndex(c => c.id === importedCustomer.id);
            if (existingIndex > -1) {
                customers[existingIndex] = importedCustomer; // 上書き
            }
            else {
                customers.push(importedCustomer);
            }
        });
        saveData();
        res.status(200).json({ message: 'Customers imported successfully.', importedCount: customersToImport.length });
    }
    catch (error) {
        console.error('Error importing customers:', error);
        res.status(500).json({ message: 'Failed to import customers.', error: error instanceof Error ? error.message : String(error) });
    }
    finally {
        // アップロードされた一時ファイルを削除
        fs_1.default.unlink(filePath, (err) => {
            if (err)
                console.error('Error deleting temporary file:', err);
        });
    }
}));
app.post('/api/import/products', upload.single('file'), (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    if (!req.file) {
        return res.status(400).json({ message: 'No file uploaded.' });
    }
    const filePath = req.file.path;
    const productsToImport = [];
    try {
        if (req.file.mimetype === 'text/csv') {
            // CSVファイルの処理
            yield new Promise((resolve, reject) => {
                fs_1.default.createReadStream(filePath)
                    .pipe((0, csv_parser_1.default)())
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
        }
        else if (req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            // Excelファイルの処理
            const workbook = new exceljs_1.default.Workbook();
            yield workbook.xlsx.readFile(filePath);
            const worksheet = workbook.getWorksheet(1);
            if (worksheet) {
                const headerMap = {};
                worksheet.getRow(1).eachCell((cell, colNumber) => {
                    headerMap[colNumber] = cell.text.trim();
                });
                worksheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1)
                        return; // Skip header row
                    const productData = {};
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
        }
        else {
            return res.status(400).json({ message: 'Unsupported file type.' });
        }
        // 既存の商品データとマージ（重複は上書き）
        productsToImport.forEach(importedProduct => {
            const existingIndex = products.findIndex(p => p.id === importedProduct.id);
            if (existingIndex > -1) {
                products[existingIndex] = importedProduct; // 上書き
            }
            else {
                products.push(importedProduct);
            }
        });
        saveData();
        res.status(200).json({ message: 'Products imported successfully.', importedCount: productsToImport.length });
    }
    catch (error) {
        console.error('Error importing products:', error);
        res.status(500).json({ message: 'Failed to import products.', error: error instanceof Error ? error.message : String(error) });
    }
    finally {
        // アップロードされた一時ファイルを削除
        fs_1.default.unlink(filePath, (err) => {
            if (err)
                console.error('Error deleting temporary file:', err);
        });
    }
}));
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
        products[productIndex] = Object.assign(Object.assign({}, products[productIndex]), updatedProductData);
        saveData();
        res.status(200).json(products[productIndex]);
    }
    else {
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
                }
                else {
                    return res.status(400).json({ message: `Customer not found: ${updatedDeliveryData.customerName}` });
                }
            }
            else if (updatedDeliveryData.customerId) {
                updatedCustomerId = updatedDeliveryData.customerId;
            }
            const updatedItems = updatedDeliveryData.items ? updatedDeliveryData.items.map((item) => {
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
                    }
                    else {
                        productIdToUse = ''; // Indicate free-form product
                    }
                }
                else if (productIdToUse) {
                    const masterProduct = products.find(p => p.id === productIdToUse);
                    if (masterProduct) {
                        productNameToUse = masterProduct.name;
                        unitToUse = unitToUse || masterProduct.unit;
                        unitPriceToUse = unitPriceToUse || masterProduct.unitPrice;
                    }
                    else {
                        // productIdが見つからない場合もエラーを返す
                        return res.status(400).json({ message: `Product not found with ID: ${productIdToUse}` });
                    }
                }
                return Object.assign(Object.assign({}, item), { productId: productIdToUse, productName: productNameToUse, unit: unitToUse, unitPrice: unitPriceToUse });
            }) : currentDelivery.items;
            deliveries[deliveryIndex] = Object.assign(Object.assign(Object.assign({}, currentDelivery), updatedDeliveryData), { items: updatedItems, customerId: updatedCustomerId });
            saveData();
            res.status(200).json(deliveries[deliveryIndex]);
        }
        catch (error) {
            console.error('Error updating delivery:', error);
            res.status(500).json({ message: 'Failed to update delivery.', error: error instanceof Error ? error.message : String(error) });
        }
    }
    else {
        res.status(404).json({ message: 'Delivery not found.' });
    }
});
app.delete('/api/products/:id', (req, res) => {
    const { id } = req.params;
    const initialLength = products.length;
    products = products.filter(product => product.id !== id);
    if (products.length < initialLength) {
        saveData();
        res.status(200).json({ message: 'Product deleted successfully.' });
    }
    else {
        res.status(404).json({ message: 'Product not found.' });
    }
});
app.delete('/api/deliveries/:id', (req, res) => {
    const { id } = req.params;
    const initialLength = deliveries.length;
    deliveries = deliveries.filter(delivery => delivery.id !== id);
    if (deliveries.length < initialLength) {
        saveData();
        res.status(200).json({ message: 'Delivery deleted successfully.' });
    }
    else {
        res.status(404).json({ message: 'Delivery not found.' });
    }
});
app.delete('/api/deliveries/:id', (req, res) => {
    const { id } = req.params;
    const initialLength = deliveries.length;
    deliveries = deliveries.filter(delivery => delivery.id !== id);
    if (deliveries.length < initialLength) {
        saveData();
        res.status(200).json({ message: 'Delivery deleted successfully.' });
    }
    else {
        res.status(404).json({ message: 'Delivery not found.' });
    }
});
// Reset Endpoints
app.delete('/api/reset/products', (req, res) => {
    products = [];
    try {
        fs_1.default.unlinkSync(productsFilePath);
    }
    catch (error) {
        console.error('Error deleting products.json:', error);
    }
    saveData();
    res.status(200).json({ message: 'Products data reset successfully.' });
});
app.delete('/api/reset/customers', (req, res) => {
    customers = [];
    try {
        fs_1.default.unlinkSync(customersFilePath);
    }
    catch (error) {
        console.error('Error deleting customers.json:', error);
    }
    saveData();
    res.status(200).json({ message: 'Customers data reset successfully.' });
});
app.delete('/api/reset/deliveries', (req, res) => {
    deliveries = [];
    try {
        fs_1.default.unlinkSync(deliveriesFilePath);
    }
    catch (error) {
        console.error('Error deleting deliveries.json:', error);
    }
    saveData();
    res.status(200).json({ message: 'Deliveries data reset successfully.' });
});
app.delete('/api/reset/invoices', (req, res) => {
    invoices = []; // 請求書にはモックデータがないため空にする
    try {
        fs_1.default.unlinkSync(invoicesFilePath);
    }
    catch (error) {
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
app.post('/api/import/deliveries', upload.single('file'), (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    if (!req.file) {
        return res.status(400).json({ message: 'No file uploaded.' });
    }
    const filePath = req.file.path;
    const deliveriesMap = {};
    try {
        if (req.file.mimetype === 'text/csv') {
            // CSVファイルの処理は現状維持（必要であれば同様のロジックを実装）
            return res.status(501).json({ message: 'CSV import for deliveries is not fully implemented with aggregation.' });
        }
        else if (req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            const workbook = new exceljs_1.default.Workbook();
            yield workbook.xlsx.readFile(filePath);
            const worksheet = workbook.getWorksheet(1);
            if (worksheet) {
                const headerMap = {};
                worksheet.getRow(1).eachCell((cell, colNumber) => {
                    headerMap[colNumber] = cell.text.trim();
                });
                worksheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1)
                        return; // Skip header row
                    const deliveryData = {};
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
                    }
                    else {
                        deliveriesMap[deliveryId] = {
                            id: deliveryId,
                            voucherNumber: deliveryData.voucherNumber,
                            deliveryDate: deliveryData.deliveryDate,
                            customerId: customer ? customer.id : '',
                            items: [newItem],
                            notes: deliveryData.notes,
                            orderId: deliveryData.orderId,
                            status: deliveryData.status,
                            invoiceStatus: deliveryData.invoiceStatus,
                            salesGroup: deliveryData.salesGroup,
                            shippingAddressName: deliveryData.shippingAddressName,
                            shippingPostalCode: deliveryData.shippingPostalCode,
                            shippingAddressDetail: deliveryData.shippingAddressDetail,
                        };
                    }
                });
            }
        }
        else {
            return res.status(400).json({ message: 'Unsupported file type.' });
        }
        const deliveriesToImport = Object.values(deliveriesMap);
        deliveriesToImport.forEach(importedDelivery => {
            const existingIndex = deliveries.findIndex(d => d.id === importedDelivery.id);
            if (existingIndex > -1) {
                deliveries[existingIndex] = importedDelivery; // 上書き
            }
            else {
                deliveries.push(importedDelivery);
            }
        });
        saveData();
        res.status(200).json({ message: 'Deliveries imported successfully.', importedCount: deliveriesToImport.length });
    }
    catch (error) {
        console.error('Error importing deliveries:', error);
        res.status(500).json({ message: 'Failed to import deliveries.', error: error instanceof Error ? error.message : String(error) });
    }
    finally {
        fs_1.default.unlink(filePath, (err) => {
            if (err)
                console.error('Error deleting temporary file:', err);
        });
    }
}));
app.listen(port, () => {
    loadData();
    console.log(`Backend server listening at http://localhost:${port}`);
});
