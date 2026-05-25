from datetime import datetime
from flask_login import UserMixin
from . import db


class AppSetting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    platform_name = db.Column(db.String(150), default='Repair Control', nullable=False)
    platform_language = db.Column(db.String(10), default='en', nullable=False)
    theme_primary = db.Column(db.String(20), default='#38bdf8')
    theme_accent = db.Column(db.String(20), default='#0f172a')
    support_enabled = db.Column(db.Boolean, default=False, nullable=False)
    support_whatsapp = db.Column(db.String(40))
    support_message = db.Column(db.String(255), default='Hello, I need help with the repair system.')
    support_embed_code = db.Column(db.Text)
    support_custom_css = db.Column(db.Text)
    invoice_title = db.Column(db.String(120), default='Repair Invoice')
    invoice_line_label = db.Column(db.String(120), default='Problem Description')
    qr_caption = db.Column(db.String(120), default='Scan for status')
    show_public_status_url = db.Column(db.Boolean, default=False, nullable=False)
    # Menu labels
    menu_dashboard = db.Column(db.String(80), default='Dashboard')
    menu_customers = db.Column(db.String(80), default='Customers')
    menu_parts = db.Column(db.String(80), default='Parts')
    menu_repairs = db.Column(db.String(80), default='Repairs')
    menu_buying = db.Column(db.String(80), default='GSM')
    menu_reports = db.Column(db.String(80), default='Reports')
    menu_audit_log = db.Column(db.String(80), default='Audit Log')
    menu_users = db.Column(db.String(80), default='Users')
    menu_settings = db.Column(db.String(80), default='Settings')
    menu_data_manager = db.Column(db.String(80), default='Data Manager')
    menu_profile = db.Column(db.String(80), default='Profile')
    menu_logout = db.Column(db.String(80), default='Logout')
    menu_companies = db.Column(db.String(80), default='Companies')
    menu_status_manager = db.Column(db.String(80), default='Status Manager')
    menu_payment_methods = db.Column(db.String(80), default='Payment Methods')
    menu_master_settings = db.Column(db.String(80), default='Master Settings')
    # Dashboard labels
    label_repairs_count = db.Column(db.String(80), default='Repairs')
    label_customers_count = db.Column(db.String(80), default='Customers')
    label_parts_count = db.Column(db.String(80), default='Parts')
    label_users_count = db.Column(db.String(80), default='Users')
    label_low_stock = db.Column(db.String(80), default='Low Stock')
    label_overdue_repairs = db.Column(db.String(80), default='Overdue Repairs')
    label_todays_repairs = db.Column(db.String(80), default="Today's repairs")
    label_recent_payments = db.Column(db.String(80), default='Recent payments')
    label_recent_activity = db.Column(db.String(80), default='Recent activity')
    label_recent_repairs = db.Column(db.String(80), default='Recent Repairs')
    label_search_placeholder = db.Column(db.String(160), default='Search by phone / repair ID / serial / IMEI / customer')
    label_repair_code = db.Column(db.String(80), default='Repair Code')
    label_customer = db.Column(db.String(80), default='Customer')
    label_device = db.Column(db.String(80), default='Device')
    label_status = db.Column(db.String(80), default='Status')
    label_action = db.Column(db.String(80), default='Action')
    label_open = db.Column(db.String(80), default='Open')
    gsm_bought_status_label = db.Column(db.String(80), default='BOUGHT BY SMP')
    gsm_sold_status_label = db.Column(db.String(80), default='SOLD BY SMP')


class Company(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    plan_type = db.Column(db.String(20), default='trial', nullable=False)
    unlimited_access = db.Column(db.Boolean, default=True, nullable=False)
    active_until = db.Column(db.Date)
    name = db.Column(db.String(150), nullable=False, unique=True)
    phone = db.Column(db.String(50))
    email = db.Column(db.String(120))
    address = db.Column(db.String(255))
    vat_number = db.Column(db.String(80))
    logo_path = db.Column(db.String(255), default='img/logo.png')
    invoice_terms = db.Column(db.Text, default='Goods remain property of the company until fully paid. Please verify device and accessories on collection.')
    storage_notice = db.Column(db.Text, default='Please collect the repaired device within 3 to 4 months after repair. Due to limited storage space, devices left longer may be handled under storage rules. Thank you for your understanding.')
    invoice_terms_en = db.Column(db.Text, default='Goods remain property of the company until fully paid. Please verify device and accessories on collection.')
    invoice_terms_fr = db.Column(db.Text, default='Les biens restent la propriété de l’entreprise jusqu’au paiement complet. Veuillez vérifier l’appareil et les accessoires lors du retrait.')
    invoice_terms_nl = db.Column(db.Text, default='Goederen blijven eigendom van het bedrijf tot volledige betaling. Controleer toestel en accessoires bij afhaling.')
    storage_notice_en = db.Column(db.Text, default='Please collect the repaired device within 3 to 4 months after repair. Due to limited storage space, devices left longer may be handled under storage rules. Thank you for your understanding.')
    storage_notice_fr = db.Column(db.Text, default='Veuillez récupérer l’appareil réparé dans les 3 à 4 mois suivant la réparation. En raison de l’espace de stockage limité, les appareils laissés plus longtemps peuvent être soumis aux règles de stockage. Merci de votre compréhension.')
    storage_notice_nl = db.Column(db.Text, default='Gelieve het herstelde toestel binnen 3 tot 4 maanden na de herstelling op te halen. Door beperkte opslagruimte kunnen toestellen die langer blijven liggen onder opslagregels vallen. Dank voor uw begrip.')

    gsm_invoice_title = db.Column(db.String(120), default='Ticket')
    gsm_terms_en = db.Column(db.Text, default='Goods remain property of the company until full payment is received. Please verify the device and accessories at the time of collection.')
    gsm_terms_fr = db.Column(db.Text, default='Les biens restent la propriété de l’entreprise jusqu’au paiement complet. Veuillez vérifier l’appareil et les accessoires lors du retrait.')
    gsm_terms_nl = db.Column(db.Text, default='Goederen blijven eigendom van het bedrijf tot volledige betaling is ontvangen. Controleer het apparaat en de accessoires bij het ophalen.')
    gsm_notice_en = db.Column(db.Text, default='Please collect the device within 90 days. Due to limited storage space, devices not collected within this period may be treated as abandoned.')
    gsm_notice_fr = db.Column(db.Text, default='Veuillez récupérer l’appareil dans les 90 jours. En raison de l’espace limité, tout appareil non récupéré dans ce délai peut être considéré comme abandonné.')
    gsm_notice_nl = db.Column(db.Text, default='Gelieve het toestel binnen 90 dagen op te halen. Door beperkte opslagruimte kan een toestel dat niet binnen deze periode wordt opgehaald als verlaten worden beschouwd.')

    default_language = db.Column(db.String(10), default='en')
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    print_width_in = db.Column(db.Float, default=3.82)
    print_height_in = db.Column(db.Float, default=5.80)
    print_padding_x_in = db.Column(db.Float, default=0.12)
    print_padding_y_in = db.Column(db.Float, default=0.12)
    print_header_font_px = db.Column(db.Integer, default=10)
    print_body_font_px = db.Column(db.Integer, default=11)
    print_notice_font_px = db.Column(db.Integer, default=9)
    print_content_shift_in = db.Column(db.Float, default=0.18)
    print_80mm_width_mm = db.Column(db.Float, default=80)
    print_80mm_padding_mm = db.Column(db.Float, default=4)
    print_80mm_header_font_px = db.Column(db.Integer, default=9)
    print_80mm_body_font_px = db.Column(db.Integer, default=9)
    print_80mm_notice_font_px = db.Column(db.Integer, default=8)
    print_80mm_content_shift_mm = db.Column(db.Float, default=0)


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=True)
    company = db.relationship('Company', backref='users')
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(30), nullable=False, default='Staff')
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    unlimited_access = db.Column(db.Boolean, default=True, nullable=False)
    active_until = db.Column(db.Date)

    @property
    def is_super_admin(self):
        return self.role == 'Super Admin'


class StatusOption(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False, unique=True)
    sort_order = db.Column(db.Integer, default=0, nullable=False)
    is_active = db.Column(db.Boolean, default=True, nullable=False)


class PaymentMethod(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False, unique=True)
    sort_order = db.Column(db.Integer, default=0, nullable=False)
    is_active = db.Column(db.Boolean, default=True, nullable=False)


class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='customers')
    name = db.Column(db.String(120), nullable=False)
    phone = db.Column(db.String(50), nullable=False)
    email = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Part(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='parts')
    name = db.Column(db.String(120), nullable=False)
    compatible_model = db.Column(db.String(120))
    quantity = db.Column(db.Integer, default=0)
    selling_price = db.Column(db.Float, default=0)
    low_stock_limit = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class RepairJob(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='repair_jobs')
    repair_code = db.Column(db.String(20), nullable=False, unique=True)
    invoice_number = db.Column(db.String(30), nullable=False, unique=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False)
    customer = db.relationship('Customer', backref='repair_jobs')
    device_brand = db.Column(db.String(80), nullable=False)
    device_model = db.Column(db.String(120), nullable=False)
    serial_number = db.Column(db.String(120))
    imei = db.Column(db.String(120))
    issue_description = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(50), nullable=False, default='Received')
    estimated_cost = db.Column(db.Float, default=0)
    deposit_amount = db.Column(db.Float, default=0)
    due_date = db.Column(db.Date)
    customer_note = db.Column(db.Text)
    is_archived = db.Column(db.Boolean, default=False, nullable=False)
    archive_reason = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items_json = db.Column(db.Text)


class RepairUpdate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='repair_updates')
    repair_job_id = db.Column(db.Integer, db.ForeignKey('repair_job.id'), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    repair_job = db.relationship('RepairJob', backref='updates')
    user = db.relationship('User')
    old_status = db.Column(db.String(50))
    new_status = db.Column(db.String(50), nullable=False)
    note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class InternalNote(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    repair_job_id = db.Column(db.Integer, db.ForeignKey('repair_job.id'), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    company = db.relationship('Company', backref='internal_notes')
    repair_job = db.relationship('RepairJob', backref='internal_notes')
    user = db.relationship('User')
    note = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class RepairPartUsed(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='parts_used_records')
    repair_job_id = db.Column(db.Integer, db.ForeignKey('repair_job.id'), nullable=False)
    part_id = db.Column(db.Integer, db.ForeignKey('part.id'), nullable=False)
    repair_job = db.relationship('RepairJob', backref='parts_used')
    part = db.relationship('Part')
    quantity = db.Column(db.Integer, nullable=False)
    price_used = db.Column(db.Float, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Payment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='payments')
    repair_job_id = db.Column(db.Integer, db.ForeignKey('repair_job.id'), nullable=False)
    repair_job = db.relationship('RepairJob', backref='payments')
    amount = db.Column(db.Float, nullable=False)
    method = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class BuyingTicket(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='buying_tickets')
    ticket_code = db.Column(db.String(20), nullable=False, unique=True)
    invoice_number = db.Column(db.String(30), nullable=False, unique=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False)
    customer = db.relationship('Customer', backref='buying_tickets')
    device_brand = db.Column(db.String(80), nullable=False)
    device_model = db.Column(db.String(120), nullable=False)
    serial_number = db.Column(db.String(120))
    imei = db.Column(db.String(120))
    issue_description = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(50), nullable=False, default='Buying')
    battery_percentage = db.Column(db.String(50))
    storage_gb = db.Column(db.String(20))
    estimated_cost = db.Column(db.Float, default=0)
    deposit_amount = db.Column(db.Float, default=0)
    due_date = db.Column(db.Date)
    customer_note = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    items_json = db.Column(db.Text)


class BuyingPayment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=False, index=True)
    company = db.relationship('Company', backref='buying_payments')
    buying_ticket_id = db.Column(db.Integer, db.ForeignKey('buying_ticket.id'), nullable=False)
    buying_ticket = db.relationship('BuyingTicket', backref='payments')
    amount = db.Column(db.Float, nullable=False)
    method = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('company.id'), nullable=True, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    action = db.Column(db.String(120), nullable=False)
    target_type = db.Column(db.String(80))
    target_id = db.Column(db.Integer)
    detail = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    company = db.relationship('Company', backref='audit_logs')
    user = db.relationship('User')
