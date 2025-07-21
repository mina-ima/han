export interface Product {
  id: string;
  name: string;
  unitPrice: number;
  unit: string; // 単位を追加
  shippingAddress: string;
  postalCode: string; // 郵便番号
  customer: string; // 取引先ID
  notes: string;
}

export interface Customer {
  id: string;
  name: string;
  address: string;
  postalCode: string; // 郵便番号
  phone: string;
  closingDay: number; // 締め日 (例: 20日)
  paymentTerms: string; // 支払条件 (例: 翌月末)
  email: string;
  contactPerson: string;
  invoiceDeliveryMethod: string; // 請求書送付方法 (例: 郵送, メール, Web)
}

export interface Delivery {
  id: number;
  deliveryDate: string;
  customerId: string;
  productId: string; // 商品ID、またはフリー入力の商品名
  quantity: number;
  unitPrice: number;
  unit: string; // 単位を追加
  notes: string;
  orderId?: string; // 注文IDを追加
  status: '発行済み' | '未発行'; // 納品書ステータスを追加
  invoiceStatus: '未請求' | '請求済み'; // 請求書ステータスを追加
  salesGroup?: string; // 売上グループを追加
  shippingAddressName?: string; // 配送先名を追加
  shippingPostalCode?: string; // 配送先郵便番号を追加
  shippingAddressDetail?: string; // 配送先住所詳細を追加
}

export interface Invoice {
  id: number;
  invoiceDate: string;
  customer: string;
  details: { deliveryId: number; productName: string; quantity: number; unitPrice: number; unit: string; amount: number; }[]; // 単位を追加
  totalAmount: number;
}

export interface User {
  id: string;
  username: string;
  email: string;
  role: '管理者' | '一般';
}

// 仮の商品データ
export const products: Product[] = [
  { id: 'P001', name: '商品A', unitPrice: 1000, unit: '個', shippingAddress: '東京都', postalCode: '100-0001', customer: 'C001', notes: '' },
  { id: 'P002', name: '商品B', unitPrice: 2500, unit: 'セット', shippingAddress: '大阪府', postalCode: '530-0001', customer: 'C002', notes: '' },
  { id: 'P003', name: '商品C', unitPrice: 500, unit: '本', shippingAddress: '福岡県', postalCode: '810-0001', customer: 'C001', notes: '' },
  { id: 'P004', name: 'ABCDEFGHIJKLMNOPQRSTU=VWXYZ', unitPrice: 12345, unit: '個', shippingAddress: '東京都', postalCode: '100-0001', customer: 'C001', notes: '' },
];

// 仮の取引先データ
export const customers: Customer[] = [
  { id: 'C001', name: '株式会社X', address: '東京都', postalCode: '100-0001', phone: '03-1111-2222', closingDay: 20, paymentTerms: '翌月末', email: 'x@example.com', contactPerson: '山田太郎', invoiceDeliveryMethod: 'メール' },
  { id: 'C002', name: '有限会社Y', address: '大阪府', postalCode: '530-0001', phone: '06-3333-4444', closingDay: 15, paymentTerms: '翌々月10日', email: 'y@example.com', contactPerson: '田中花子', invoiceDeliveryMethod: '郵送' },
  { id: 'C003', name: '合同会社Z', address: '福岡県', postalCode: '810-0001', phone: '092-555-6666', closingDay: 30, paymentTerms: '当月末', email: 'z@example.com', contactPerson: '鈴木一郎', invoiceDeliveryMethod: 'Web' },
  { id: 'C004', name: '赤サタな濱家らわわジェイ着', address: '東京都', postalCode: '100-0001', phone: '03-9999-8888', closingDay: 25, paymentTerms: '翌月末', email: 'test@example.com', contactPerson: 'テスト太郎', invoiceDeliveryMethod: 'メール' },
];

// 仮の納品データ
export const mockDeliveries: Delivery[] = [
  { id: 1, deliveryDate: '2024-07-01', customerId: 'C001', productId: 'P001', quantity: 10, unitPrice: 1000, unit: '個', notes: '', orderId: 'L001', status: '発行済み', salesGroup: 'SG-2024-001', invoiceStatus: '未請求', shippingAddressName: '株式会社X', shippingPostalCode: '100-0001', shippingAddressDetail: '東京都千代田区1-1-1' },
  { id: 2, deliveryDate: '2024-07-05', customerId: 'C002', productId: 'P002', quantity: 5, unitPrice: 2500, unit: 'セット', notes: '急ぎ', orderId: 'L002', status: '未発行', salesGroup: 'SG-2024-002', invoiceStatus: '未請求', shippingAddressName: '有限会社Y', shippingPostalCode: '530-0001', shippingAddressDetail: '大阪府大阪市北区2-2-2' },
  { id: 3, deliveryDate: '2024-07-10', customerId: 'C001', productId: 'P003', quantity: 20, unitPrice: 500, unit: '本', notes: '', orderId: 'L003', status: '発行済み', salesGroup: 'SG-2024-001', invoiceStatus: '未請求', shippingAddressName: '株式会社X', shippingPostalCode: '100-0001', shippingAddressDetail: '東京都千代田区1-1-1' },
  { id: 4, deliveryDate: '2024-07-12', customerId: 'C003', productId: 'P001', quantity: 8, unitPrice: 1000, unit: '個', notes: '', orderId: 'L004', status: '未発行', salesGroup: 'SG-2024-003', invoiceStatus: '未請求', shippingAddressName: '合同会社Z', shippingPostalCode: '810-0001', shippingAddressDetail: '福岡県福岡市中央区3-3-3' },
  { id: 5, deliveryDate: '2024-07-15', customerId: 'C002', productId: 'P003', quantity: 15, unitPrice: 500, unit: '本', notes: '', orderId: 'L005', status: '未発行', salesGroup: 'SG-2024-002', invoiceStatus: '未請求', shippingAddressName: '有限会社Y', shippingPostalCode: '530-0001', shippingAddressDetail: '大阪府大阪市北区2-2-2' },
  { id: 6, deliveryDate: '2024-07-17', customerId: 'C004', productId: 'P004', quantity: 1, unitPrice: 12345, unit: '個', notes: '長い名前のテスト', orderId: 'L006', status: '未発行', salesGroup: 'SG-2024-004', invoiceStatus: '未請求', shippingAddressName: '赤サタな濱家らわわジェイ着', shippingPostalCode: '100-0001', shippingAddressDetail: '東京都千代田区4-4-4' },
];

// 仮のユーザーデータ
export const users: User[] = [
  { id: 'U001', username: 'admin', email: 'admin@example.com', role: '管理者' },
  { id: 'U002', username: 'user1', email: 'user1@example.com', role: '一般' },
];

// 仮の請求書データ
export const invoices: Invoice[] = [];