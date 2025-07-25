export interface Product {
  id: string;
  name: string;
  unitPrice: number;
  unit: string; // 単位を追加
  shippingAddress: string;
  postalCode: string; // 郵便番号
  customer: string; // 取引先ID
  notes: string;
  shippingName?: string;
}

export interface Customer {
  id: string;
  name: string;
  formalName?: string; // 正式名称を追加
  address: string;
  postalCode: string; // 郵便番号
  phone: string;
  closingDay: number; // 締め日 (例: 20日)
  paymentTerms: string; // 支払条件 (例: 翌月末)
  email: string;
  contactPerson: string;
  invoiceDeliveryMethod: string; // 請求書送付方法 (例: 郵送, メール, Web)
}

export interface DeliveryItem {
  productId?: string; // Optional, for free-form input
  productName?: string; // Added for free-form input
  quantity: number;
  unitPrice: number;
  unit: string;
  notes?: string;
}

export interface Delivery {
  id: string;
  voucherNumber: string;
  deliveryDate: string;
  customerId: string;
  items: DeliveryItem[];
  notes?: string;
  orderId?: string;
  status?: '発行済み' | '未発行';
  invoiceStatus?: '未請求' | '請求済み';
  salesGroup?: string;
  shippingAddressName?: string;
  shippingPostalCode?: string;
  shippingAddressDetail?: string;
}

export interface User {
  id: string;
  username: string;
  email: string;
  role: '管理者' | '一般';
}

export interface CompanyInfo {
  name: string;
  postalCode: string;
  address: string;
  phone: string;
  fax: string;
  bankName: string;
  bankBranch: string;
  bankAccountType: string;
  bankAccountNumber: string;
  bankAccountHolder: string;
  contactPerson: string;
}
