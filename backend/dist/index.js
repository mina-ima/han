"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const masterData_1 = require("./data/masterData");
const json2csv = require('json-2-csv').json2csv; // 型定義の問題を回避
const app = (0, express_1.default)();
const port = 3002;
app.use((0, cors_1.default)({
    origin: 'http://localhost:3000',
}));
// Helper function to send data as CSV
const sendCsv = (res, data, filename) => {
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
    const { productName, productName_matchType, unit, unit_matchType, postalCode, postalCode_matchType, shippingAddress, shippingAddress_matchType, customer, customer_matchType, notes, notes_matchType, minUnitPrice, maxUnitPrice } = query;
    let filteredProducts = masterData_1.products;
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
    const { startDate, endDate, customerId, productId, minQuantity, maxQuantity, minUnitPrice, maxUnitPrice, status, salesGroup, unit, orderId, notes, minAmount, maxAmount, invoiceStatus } = query;
    let filteredDeliveries = masterData_1.mockDeliveries;
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
        filteredDeliveries = filteredDeliveries.filter(d => d.productId === productId);
    }
    if (minQuantity) {
        filteredDeliveries = filteredDeliveries.filter(d => d.quantity >= parseFloat(minQuantity));
    }
    if (maxQuantity) {
        filteredDeliveries = filteredDeliveries.filter(d => d.quantity <= parseFloat(maxQuantity));
    }
    if (minUnitPrice) {
        filteredDeliveries = filteredDeliveries.filter(d => d.unitPrice >= parseFloat(minUnitPrice));
    }
    if (maxUnitPrice) {
        filteredDeliveries = filteredDeliveries.filter(d => d.unitPrice <= parseFloat(maxUnitPrice));
    }
    if (status) {
        filteredDeliveries = filteredDeliveries.filter(d => d.status === status);
    }
    if (salesGroup) {
        filteredDeliveries = filteredDeliveries.filter(d => d.salesGroup && d.salesGroup.includes(salesGroup));
    }
    if (unit) {
        filteredDeliveries = filteredDeliveries.filter(d => d.unit && d.unit.includes(unit));
    }
    if (orderId) {
        filteredDeliveries = filteredDeliveries.filter(d => d.orderId && d.orderId.includes(orderId));
    }
    if (notes) {
        filteredDeliveries = filteredDeliveries.filter(d => d.notes && d.notes.includes(notes));
    }
    if (minAmount) {
        filteredDeliveries = filteredDeliveries.filter(d => (d.quantity * d.unitPrice) >= parseFloat(minAmount));
    }
    if (maxAmount) {
        filteredDeliveries = filteredDeliveries.filter(d => (d.quantity * d.unitPrice) <= parseFloat(maxAmount));
    }
    if (invoiceStatus) {
        filteredDeliveries = filteredDeliveries.filter(d => d.invoiceStatus === invoiceStatus);
    }
    return filteredDeliveries;
};
// Helper function to filter customers
const filterCustomers = (query) => {
    const { name, name_matchType, postalCode, postalCode_matchType, address, address_matchType, phone, phone_matchType, paymentTerms, paymentTerms_matchType, email, email_matchType, contactPerson, contactPerson_matchType, minClosingDay, maxClosingDay, invoiceDeliveryMethod } = query;
    let filteredCustomers = masterData_1.customers;
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
    let filteredUsers = masterData_1.users;
    if (username) {
        const matchType = username_matchType || 'partial';
        if (matchType === 'exact') {
            filteredUsers = filteredUsers.filter(u => u.username === username);
        }
        else { // partial
            filteredUsers = filteredUsers.filter(u => u.username.includes(username));
        }
    }
    if (email) {
        const matchType = email_matchType || 'partial';
        if (matchType === 'exact') {
            filteredUsers = filteredUsers.filter(u => u.email === email);
        }
        else { // partial
            filteredUsers = filteredUsers.filter(u => u.email.includes(email));
        }
    }
    if (role) {
        filteredUsers = filteredUsers.filter(u => u.role === role);
    }
    return filteredUsers;
};
// CSV Export Endpoints
app.get('/api/export/products', (req, res) => {
    const filteredProducts = filterProducts(req.query);
    sendCsv(res, filteredProducts, 'products.csv');
});
app.get('/api/export/customers', (req, res) => {
    const filteredCustomers = filterCustomers(req.query);
    sendCsv(res, filteredCustomers, 'customers.csv');
});
app.get('/api/export/deliveries', (req, res) => {
    const filteredDeliveries = filterDeliveries(req.query);
    sendCsv(res, filteredDeliveries, 'deliveries.csv');
});
app.get('/api/export/users', (req, res) => {
    const filteredUsers = filterUsers(req.query);
    sendCsv(res, filteredUsers, 'users.csv');
});
app.get('/api/export/salesSummary', (req, res) => {
    const salesByCustomer = {};
    masterData_1.mockDeliveries.forEach(delivery => {
        var _a;
        const customerName = ((_a = masterData_1.customers.find(c => c.id === delivery.customerId)) === null || _a === void 0 ? void 0 : _a.name) || '不明';
        const amount = delivery.quantity * delivery.unitPrice;
        salesByCustomer[customerName] = (salesByCustomer[customerName] || 0) + amount;
    });
    const dataToExport = Object.keys(salesByCustomer).map(customerName => ({
        customerName,
        totalSales: salesByCustomer[customerName],
    }));
    sendCsv(res, dataToExport, 'sales_summary.csv');
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
    const salesByCustomer = {};
    filteredDeliveries.forEach(delivery => {
        var _a;
        const customerName = ((_a = masterData_1.customers.find(c => c.id === delivery.customerId)) === null || _a === void 0 ? void 0 : _a.name) || '不明';
        const amount = delivery.quantity * delivery.unitPrice;
        salesByCustomer[customerName] = (salesByCustomer[customerName] || 0) + amount;
    });
    const dataToExport = Object.keys(salesByCustomer).map(customerName => ({
        customerName,
        totalSales: salesByCustomer[customerName],
    }));
    res.json({ summary: dataToExport, details: filteredDeliveries });
});
app.listen(port, () => {
    console.log(`Backend server listening at http://localhost:${port}`);
});
