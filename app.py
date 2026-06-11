import os
import sys
import ast
import json
import io
from datetime import datetime, date
import calendar
from sqlalchemy import text
from sqlalchemy.pool import NullPool

# Python 3.14 compatibility shim for Werkzeug/Flask internals
if sys.version_info >= (3, 14) and not hasattr(ast, 'Str'):
    class Py314Str(ast.Constant):
        def __init__(self, s=None, *args, **kwargs):
            super().__init__(value=s, kind=None, *args, **kwargs)

        @property
        def s(self):
            return self.value

    ast.Str = Py314Str

from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

import sys
from export_excel import (
    export_purchases, export_sales, export_products, 
    export_suppliers, export_customers, export_payments, 
    export_bills, excel_to_bytes
)
# When bundled by PyInstaller, resources are available in sys._MEIPASS
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

template_folder = os.path.join(base_path, 'templates')
static_folder = os.path.join(base_path, 'static')

# Provide an explicit `instance_path` to avoid Flask's auto package discovery
# which uses `pkgutil.get_loader` (removed in Python 3.14).
app = Flask(__name__, template_folder=template_folder, static_folder=static_folder, instance_path=base_path, root_path=base_path)
# configure persistent database path
env_db = os.environ.get('DATABASE_URL')
if env_db:
    app.config['SQLALCHEMY_DATABASE_URI'] = env_db
else:
    if getattr(sys, 'frozen', False):
        # When bundled as an EXE, place the DB in a persistent user folder
        data_dir = os.path.join(os.path.expanduser('~'), 'MediPharma')
        os.makedirs(data_dir, exist_ok=True)
        db_path = os.path.join(data_dir, 'billing.db')
    else:
        # during development keep DB next to the source
        db_path = os.path.join(base_path, 'billing.db')
    app.config['SQLALCHEMY_DATABASE_URI'] = f"sqlite:///{db_path}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
# Use a NullPool when running under pytest to ensure connections are
# closed immediately and avoid Windows file-lock issues on temporary DBs.
running_under_test = 'PYTEST_CURRENT_TEST' in os.environ or 'pytest' in sys.modules
if running_under_test:
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'connect_args': {'timeout': 30, 'check_same_thread': False},
        'poolclass': NullPool,
        'echo': False,
    }
else:
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'connect_args': {'timeout': 30, 'check_same_thread': False},
        'pool_pre_ping': True,  # Verify connections before using
        'pool_recycle': 3600,   # Recycle connections after 1 hour
        'echo': False,
    }
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'change-this-secret')

db = SQLAlchemy(app)


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(50), default='cashier')

    def set_password(self, pwd):
        self.password_hash = generate_password_hash(pwd)

    def check_password(self, pwd):
        return check_password_hash(self.password_hash, pwd)


class ShopProfile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False, default='MediPharma')
    gst = db.Column(db.String(50), nullable=True)
    address = db.Column(db.String(500), nullable=True)
    phone = db.Column(db.String(50), nullable=True)
    email = db.Column(db.String(100), nullable=True)


class Supplier(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    phone = db.Column(db.String(50))
    gst = db.Column(db.String(50))


class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    phone = db.Column(db.String(50))
    address = db.Column(db.String(500), nullable=True)
    payment_due = db.Column(db.Float, default=0.0)
    payment_complete = db.Column(db.Boolean, default=False)
    remark = db.Column(db.String(500), nullable=True)


class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    sku = db.Column(db.String(80), unique=True)
    quantity = db.Column(db.Integer, default=0)
    price = db.Column(db.Float, default=0.0)  # selling price (per unit or per strip)
    purchase_price = db.Column(db.Float, default=0.0)
    mrp_price = db.Column(db.Float, default=0.0)
    gst_percent = db.Column(db.Float, default=0.0)
    medicine_type = db.Column(db.String(50), default='LIQUID')
    tablets_per_strip = db.Column(db.Integer, default=0)
    expiry_date = db.Column(db.Date, nullable=True)

    @property
    def is_strip(self):
        return (self.medicine_type or '').strip().upper() == 'STRIPS' and self.tablets_per_strip > 0

    @property
    def strip_stock(self):
        if self.is_strip:
            return self.quantity // self.tablets_per_strip
        return 0

    @property
    def remaining_tablets(self):
        if self.is_strip:
            return self.quantity % self.tablets_per_strip
        return self.quantity

    @property
    def stock_display(self):
        if self.is_strip:
            strips = self.strip_stock
            tablets = self.remaining_tablets
            return f"{strips} strips + {tablets} tablets" if strips or tablets else '0'
        return str(self.quantity)


class Purchase(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey('supplier.id'))
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'))
    sku = db.Column(db.String(80), nullable=True)
    qty = db.Column(db.Integer, nullable=False)
    price = db.Column(db.Float, nullable=False)
    mrp_price = db.Column(db.Float, default=0.0)
    gst_percent = db.Column(db.Float, default=0.0)
    cgst_percent = db.Column(db.Float, default=0.0)
    sgst_percent = db.Column(db.Float, default=0.0)
    discount = db.Column(db.Float, default=0.0)
    total = db.Column(db.Float, nullable=False)
    received_by = db.Column(db.String(200), nullable=True)
    date = db.Column(db.Date, default=date.today)
    due_date = db.Column(db.Date, nullable=True)


class Sale(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'))
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'))
    sku = db.Column(db.String(80), nullable=True)
    qty = db.Column(db.Integer, nullable=False)
    price = db.Column(db.Float, nullable=False)
    mrp_price = db.Column(db.Float, default=0.0)
    gst_percent = db.Column(db.Float, default=0.0)
    cgst_percent = db.Column(db.Float, default=0.0)
    sgst_percent = db.Column(db.Float, default=0.0)
    discount = db.Column(db.Float, default=0.0)
    total = db.Column(db.Float, nullable=False)
    seller_name = db.Column(db.String(200), nullable=True)
    date = db.Column(db.Date, default=date.today)
    due_date = db.Column(db.Date, nullable=True)


class PurchaseReturn(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey('supplier.id'))
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'))
    sku = db.Column(db.String(80), nullable=True)
    qty = db.Column(db.Integer, nullable=False)
    price = db.Column(db.Float, nullable=False)
    total = db.Column(db.Float, nullable=False)
    remark = db.Column(db.String(500), nullable=True)
    date = db.Column(db.Date, default=date.today)


class SaleReturn(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'))
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'))
    sku = db.Column(db.String(80), nullable=True)
    qty = db.Column(db.Integer, nullable=False)
    price = db.Column(db.Float, nullable=False)
    total = db.Column(db.Float, nullable=False)
    remark = db.Column(db.String(500), nullable=True)
    date = db.Column(db.Date, default=date.today)


class Payment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    party_type = db.Column(db.String(20))  # 'customer' or 'supplier'
    party_id = db.Column(db.Integer)
    amount = db.Column(db.Float, nullable=False)
    mode = db.Column(db.String(20))  # full, partial, none
    payment_method = db.Column(db.String(50), nullable=True)
    payment_time = db.Column(db.Time, nullable=True)
    direction = db.Column(db.String(10))  # 'in' (customer paid) or 'out' (paid to supplier)
    date = db.Column(db.Date, default=date.today)


class ScheduledPayment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    party_type = db.Column(db.String(20))  # 'customer' or 'supplier'
    party_id = db.Column(db.Integer)
    amount_due = db.Column(db.Float, nullable=False)
    due_date = db.Column(db.Date, nullable=True)
    direction = db.Column(db.String(10))  # 'in' or 'out'
    created = db.Column(db.Date, default=date.today)
    status = db.Column(db.String(20), default='pending')  # pending, paid, cancelled


class Reminder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    scheduled_payment_id = db.Column(db.Integer, db.ForeignKey('scheduled_payment.id'))
    message = db.Column(db.String(500))
    date = db.Column(db.Date, default=date.today)
    sent = db.Column(db.Boolean, default=False)


class Notification(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    title = db.Column(db.String(200), nullable=False)
    message = db.Column(db.String(500), nullable=False)
    notification_type = db.Column(db.String(50), default='info')  # info, success, warning, danger, payment_due, low_stock
    related_id = db.Column(db.Integer, nullable=True)
    read = db.Column(db.Boolean, default=False)
    created = db.Column(db.DateTime, default=datetime.now)
    action_url = db.Column(db.String(500), nullable=True)


def create_default_admin():
    db.session.rollback()
    try:
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            admin = User(username='admin', role='admin')
            admin.set_password(os.environ.get('DEFAULT_ADMIN_PASS', 'admin1472'))
            db.session.add(admin)
            db.session.commit()
        return admin
    except Exception:
        db.session.rollback()
        return User.query.filter_by(username='admin').first()


def create_default_shop_profile():
    db.session.rollback()
    shop = ShopProfile.query.first()
    if not shop:
        shop = ShopProfile(name='MediPharma', gst='', address='', phone='', email='')
        db.session.add(shop)
        db.session.commit()
    return shop


def ensure_db_column(table_name, column_name, column_definition):
    try:
        result = db.session.execute(text(f"PRAGMA table_info('{table_name}')")).all()
        columns = [row[1] for row in result]
        if column_name not in columns:
            db.session.execute(text(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_definition}"))
            db.session.commit()
    except Exception:
        pass


def upgrade_database_schema():
    ensure_db_column('customer', 'address', 'VARCHAR(500)')
    ensure_db_column('customer', 'payment_due', 'FLOAT DEFAULT 0.0')
    ensure_db_column('customer', 'payment_complete', 'BOOLEAN DEFAULT 0')
    ensure_db_column('customer', 'remark', 'VARCHAR(500)')
    ensure_db_column('product', 'purchase_price', 'FLOAT DEFAULT 0.0')
    ensure_db_column('product', 'mrp_price', 'FLOAT DEFAULT 0.0')
    ensure_db_column('product', 'medicine_type', 'VARCHAR(50) DEFAULT "LIQUID"')
    ensure_db_column('product', 'tablets_per_strip', 'INTEGER DEFAULT 0')
    ensure_db_column('purchase', 'gst_percent', 'FLOAT DEFAULT 0.0')
    ensure_db_column('purchase', 'cgst_percent', 'FLOAT DEFAULT 0.0')
    ensure_db_column('purchase', 'sgst_percent', 'FLOAT DEFAULT 0.0')
    ensure_db_column('purchase', 'discount', 'FLOAT DEFAULT 0.0')
    ensure_db_column('purchase', 'due_date', 'DATE')
    ensure_db_column('purchase', 'received_by', 'VARCHAR(200)')
    ensure_db_column('purchase', 'mrp_price', 'FLOAT DEFAULT 0.0')
    ensure_db_column('purchase', 'sku', 'VARCHAR(80)')
    ensure_db_column('sale', 'gst_percent', 'FLOAT DEFAULT 0.0')
    ensure_db_column('sale', 'cgst_percent', 'FLOAT DEFAULT 0.0')
    ensure_db_column('sale', 'sgst_percent', 'FLOAT DEFAULT 0.0')
    ensure_db_column('sale', 'discount', 'FLOAT DEFAULT 0.0')
    ensure_db_column('sale', 'seller_name', 'VARCHAR(200)')
    ensure_db_column('sale', 'due_date', 'DATE')
    ensure_db_column('sale', 'mrp_price', 'FLOAT DEFAULT 0.0')
    ensure_db_column('sale', 'sku', 'VARCHAR(80)')
    ensure_db_column('payment', 'payment_method', 'VARCHAR(50)')
    ensure_db_column('payment', 'payment_time', 'TIME')


@app.context_processor
def inject_shop_profile():
    shop = ShopProfile.query.first()
    return {'shop_profile': shop}


def verify_database_integrity():
    """Verify and repair database schema integrity"""
    try:
        # Dispose of all connections to force fresh connections
        db.engine.dispose()
        
        # Create all tables
        db.create_all()
        
        # Run schema upgrades
        upgrade_database_schema()
        
        # Verify critical columns exist
        with db.engine.connect() as conn:
            result = conn.execute(text("PRAGMA table_info(sale)"))
            columns = {row[1] for row in result}
            required = {'gst_percent', 'discount', 'seller_name', 'due_date', 'qty', 'price', 'total', 'customer_id', 'product_id', 'date'}
            missing = required - columns
            
            if missing:
                print(f"WARNING: Sale table missing columns: {missing}")
            else:
                print("✓ Sale table schema verified successfully")
        
        print("✓ Database integrity check passed")
    except Exception as e:
        print(f"Database integrity check warning: {e}")
        # Don't fail on warnings, continue startup


def init_db(start_scheduler=True):
    db.create_all()
    running_under_test = 'PYTEST_CURRENT_TEST' in os.environ or 'pytest' in sys.modules
    if not running_under_test:
        create_default_admin()
        create_default_shop_profile()
    upgrade_database_schema()
    verify_database_integrity()
    if start_scheduler:
        scheduler = BackgroundScheduler()
        scheduler.add_job(func=check_scheduled_payments, trigger='interval', days=1, next_run_time=datetime.now())
        scheduler.start()

def login_required(f):
    from functools import wraps

    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)

    return decorated


def role_required(roles):
    from functools import wraps

    def wrapper(f):
        @wraps(f)
        def inner(*args, **kwargs):
            uid = session.get('user_id')
            if not uid:
                return redirect(url_for('login'))
            user = db.session.get(User, uid)
            if not user or user.role not in roles:
                flash('Permission denied', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)

        return inner

    return wrapper


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        pin = (request.form.get('pin') or '').strip()
        access_pin = os.environ.get('ACCESS_PIN', '147369')

        if len(pin) != 6 or not pin.isdigit():
            flash('Enter a valid 6-digit PIN', 'danger')
        elif pin == access_pin:
            user = User.query.order_by(User.id).first()
            if not user:
                user = User(username='pin_admin', role='admin')
                user.set_password(access_pin)
                db.session.add(user)
                db.session.commit()
            session['user_id'] = user.id
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid PIN', 'danger')
    return render_template('billing.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/change-password', methods=['GET', 'POST'])
@login_required
def change_password():
    user = db.session.get(User, session['user_id'])

    if request.method == 'POST':
        new_username = (request.form.get('new_username') or '').strip()
        old = request.form.get('old')
        new = request.form.get('new')

        if not user:
            flash('User session expired. Please log in again.', 'danger')
            return redirect(url_for('login'))

        if new_username and new_username != user.username:
            existing_user = User.query.filter(User.username == new_username, User.id != user.id).first()
            if existing_user:
                flash('That username is already taken.', 'danger')
                return render_template('change_password.html', user=user)
            user.username = new_username
            flash('Username updated successfully.', 'success')

        if new:
            if not old:
                flash('Current password is required to change password.', 'danger')
                return render_template('change_password.html', user=user)
            if not user.check_password(old):
                flash('Current password incorrect.', 'danger')
                return render_template('change_password.html', user=user)
            user.set_password(new)
            db.session.commit()
            # after password change force re-login so new password is required
            flash('Password changed. Please login with your new password.', 'success')
            session.clear()
            return redirect(url_for('login'))

        db.session.commit()

    return render_template('change_password.html', user=user)


@app.route('/')
@login_required
def dashboard():
    products = Product.query.all()
    low_stock = [p for p in products if p.quantity <= 5]
    today = date.today()
    expiring = [p for p in products if p.expiry_date and p.expiry_date <= today]
    total_stock_value = sum(p.quantity * p.price for p in products)
    sales = Sale.query.all()
    purchases = Purchase.query.all()
    purchase_returns = PurchaseReturn.query.all()
    sale_returns = SaleReturn.query.all()
    total_sales = sum(s.total for s in sales)
    total_purchases = sum(p.total for p in purchases)
    total_purchase_returns = sum(r.total for r in purchase_returns)
    total_sale_returns = sum(r.total for r in sale_returns)
    profit = total_sales - total_purchases - total_sale_returns + total_purchase_returns
    
    # fetch notifications and pending payments
    uid = session['user_id']
    unread_notifs = Notification.query.filter_by(user_id=uid, read=False).order_by(Notification.created.desc()).limit(5).all()
    pending_payments = ScheduledPayment.query.filter_by(status='pending').limit(5).all()
    
    return render_template('dashboard.html', low_stock=low_stock, expiring=expiring,
                           total_stock_value=total_stock_value, profit=profit,
                           unread_notifs=unread_notifs, pending_payments=pending_payments)


@app.route('/inventory')
@login_required
def inventory():
    products = Product.query.all()
    return render_template('inventory.html', products=products, now=datetime.now())


@app.route('/purchase', methods=['GET', 'POST'])
@login_required
def purchase():
    suppliers = Supplier.query.all()
    products = Product.query.all()
    if request.method == 'POST':
        supplier_id = request.form.get('supplier')
        received_by = request.form.get('received_by')
        payment_mode = request.form.get('payment_mode')
        payment_method = request.form.get('payment_method')
        payment_amount = float(request.form.get('payment_amount') or 0)
        payment_time = request.form.get('payment_time')
        date_val = request.form.get('date')
        due_date_val = request.form.get('due_date')
        date_val = datetime.strptime(date_val, '%Y-%m-%d').date() if date_val else date.today()
        due_date = datetime.strptime(due_date_val, '%Y-%m-%d').date() if due_date_val else None

        product_ids = request.form.getlist('product[]')
        qtys = request.form.getlist('qty[]')
        prices = request.form.getlist('price[]')
        mrps = request.form.getlist('mrp[]')
        skus = request.form.getlist('sku[]')
        cgsts = request.form.getlist('cgst[]')
        sgsts = request.form.getlist('sgst[]')
        gsts = request.form.getlist('gst[]')
        discounts = request.form.getlist('discount[]')
        expiries = request.form.getlist('expiry_date[]')
        selling_prices = request.form.getlist('selling_price[]')

        total_amount = 0.0
        created_purchases = []
        for idx, product_id in enumerate(product_ids):
            if not product_id:
                continue
            try:
                qty = int(qtys[idx] or 0)
                price = float(prices[idx] or 0)
                mrp_price = float(mrps[idx] or 0)
                cgst = float(cgsts[idx] or 0)
                sgst = float(sgsts[idx] or 0)
                total_gst = cgst + sgst
                if total_gst == 0 and idx < len(gsts):
                    total_gst = float(gsts[idx] or 0)
                    cgst = total_gst / 2.0
                    sgst = total_gst / 2.0
                discount = float(discounts[idx] or 0)
            except Exception:
                continue
            if qty <= 0 or price < 0:
                continue
            # Per-user request: calculate CGST/SGST based on product purchase price when available
            prod = Product.query.get(product_id)
            base_unit_price = None
            if prod and prod.purchase_price and prod.purchase_price > 0:
                base_unit_price = prod.purchase_price
            else:
                base_unit_price = price

            subtotal = qty * base_unit_price
            gst_amount = subtotal * total_gst / 100.0
            total = subtotal + gst_amount - discount
            total_amount += total
            sku_value = skus[idx] if idx < len(skus) else None
            purchase = Purchase(supplier_id=supplier_id, product_id=product_id, sku=sku_value, qty=qty, price=price,
                                mrp_price=mrp_price, gst_percent=total_gst, cgst_percent=cgst, sgst_percent=sgst,
                                discount=discount, total=total, received_by=received_by,
                                date=date_val, due_date=due_date)
            prod = Product.query.get(product_id)
            if prod:
                add_qty = qty
                if prod.is_strip and prod.tablets_per_strip > 0:
                    add_qty = qty * prod.tablets_per_strip
                prod.quantity += add_qty
                # If product is strips, store purchase_price as per-tablet price
                try:
                    if prod.is_strip and prod.tablets_per_strip and prod.tablets_per_strip > 0:
                        prod.purchase_price = float(price) / float(prod.tablets_per_strip)
                    else:
                        prod.purchase_price = price
                except Exception:
                    prod.purchase_price = price
                # Update product-level MRP and SKU if provided during purchase
                prod.mrp_price = mrp_price
                if sku_value:
                    prod.sku = sku_value
                expiry_value = expiries[idx] if idx < len(expiries) else None
                if expiry_value:
                    try:
                        if len(expiry_value) == 7 and '-' in expiry_value:
                            y, m = expiry_value.split('-')
                            y = int(y); m = int(m)
                            last_day = calendar.monthrange(y, m)[1]
                            prod.expiry_date = date(y, m, last_day)
                        else:
                            prod.expiry_date = datetime.strptime(expiry_value, '%Y-%m-%d').date()
                    except Exception:
                        pass
                selling_price = selling_prices[idx] if idx < len(selling_prices) else None
                try:
                    spv = float(selling_price or 0)
                    if spv and spv > 0:
                        prod.price = spv
                except Exception:
                    pass
            db.session.add(purchase)
            created_purchases.append(purchase)

        if not created_purchases:
            flash('Please add at least one product to record purchase.', 'danger')
            return redirect(url_for('purchase'))

        if payment_mode in ('full', 'partial') and payment_amount > 0:
            payment = Payment(party_type='supplier', party_id=supplier_id, amount=payment_amount,
                              mode=payment_mode, payment_method=payment_method,
                              payment_time=datetime.strptime(payment_time, '%H:%M').time() if payment_time else None,
                              direction='out', date=date_val)
            db.session.add(payment)

        due_amount = 0.0
        if payment_mode == 'partial':
            due_amount = max(0.0, total_amount - payment_amount)
        elif payment_mode == 'none':
            due_amount = total_amount
        elif payment_mode == 'full':
            due_amount = max(0.0, total_amount - payment_amount)

        if due_amount > 0:
            sp = ScheduledPayment(party_type='supplier', party_id=supplier_id, amount_due=due_amount,
                                  due_date=due_date, direction='out')
            db.session.add(sp)

        db.session.commit()
        for purchase in created_purchases:
            prod = Product.query.get(purchase.product_id)
            if prod:
                check_low_stock_and_alert(prod)

        flash('Purchase recorded', 'success')
        return redirect(url_for('purchase'))

    recent_purchases = Purchase.query.order_by(Purchase.date.desc(), Purchase.id.desc()).limit(100).all()
    for p in recent_purchases:
        p.supplier = Supplier.query.get(p.supplier_id) if p.supplier_id else None
        p.product = Product.query.get(p.product_id) if p.product_id else None
    return render_template('purchase.html', suppliers=suppliers, products=products, purchases=recent_purchases)


@app.route('/billing', methods=['GET', 'POST'])
@login_required
def billing():
    customers = Customer.query.all()
    products = Product.query.all()
    if request.method == 'POST':
        customer_id = request.form.get('customer')
        seller_name = request.form.get('seller_name')
        payment_mode = request.form.get('payment_mode')
        payment_method = request.form.get('payment_method')
        payment_amount = float(request.form.get('payment_amount') or 0)
        payment_time = request.form.get('payment_time')
        date_val = request.form.get('date')
        due_date_val = request.form.get('due_date')
        date_val = datetime.strptime(date_val, '%Y-%m-%d').date() if date_val else date.today()
        due_date = datetime.strptime(due_date_val, '%Y-%m-%d').date() if due_date_val else None

        product_ids = request.form.getlist('product[]')
        qtys = request.form.getlist('qty[]')
        prices = request.form.getlist('price[]')
        mrps = request.form.getlist('mrp[]')
        skus = request.form.getlist('sku[]')
        cgsts = request.form.getlist('cgst[]')
        sgsts = request.form.getlist('sgst[]')
        gsts = request.form.getlist('gst[]')
        discounts = request.form.getlist('discount[]')

        total_amount = 0.0
        created_sales = []
        for idx, product_id in enumerate(product_ids):
            if not product_id:
                continue
            try:
                qty = int(qtys[idx] or 0)
                price = float(prices[idx] or 0)
                mrp_price = float(mrps[idx] or 0)
                cgst = float(cgsts[idx] or 0)
                sgst = float(sgsts[idx] or 0)
                total_gst = cgst + sgst
                if total_gst == 0 and idx < len(gsts):
                    total_gst = float(gsts[idx] or 0)
                    cgst = total_gst / 2.0
                    sgst = total_gst / 2.0
                discount = float(discounts[idx] or 0)
            except Exception:
                continue
            if qty <= 0 or price < 0:
                continue
            subtotal = qty * price
            gst_amount = subtotal * total_gst / 100.0
            total = subtotal + gst_amount - discount
            total_amount += total
            sku_value = skus[idx] if idx < len(skus) else None
            sale = Sale(customer_id=customer_id, product_id=product_id, sku=sku_value, qty=qty, price=price,
                        mrp_price=mrp_price, gst_percent=total_gst, cgst_percent=cgst, sgst_percent=sgst,
                        discount=discount, total=total, seller_name=seller_name,
                        date=date_val, due_date=due_date)
            # Reduce inventory: stored quantity is in units (tablets) for strips
            prod = Product.query.get(product_id)
            if prod:
                prod.quantity = max(0, prod.quantity - qty)
            db.session.add(sale)
            created_sales.append(sale)

        if not created_sales:
            flash('Please add at least one product to the bill.', 'danger')
            return redirect(url_for('billing'))

        if payment_mode in ('full', 'partial') and payment_amount > 0:
            payment = Payment(party_type='customer', party_id=customer_id, amount=payment_amount,
                              mode=payment_mode, payment_method=payment_method,
                              payment_time=datetime.strptime(payment_time, '%H:%M').time() if payment_time else None,
                              direction='in', date=date_val)
            db.session.add(payment)

        due_amount = 0.0
        if payment_mode == 'partial':
            due_amount = max(0.0, total_amount - payment_amount)
        elif payment_mode == 'none':
            due_amount = total_amount
        elif payment_mode == 'full':
            due_amount = max(0.0, total_amount - payment_amount)

        if due_amount > 0:
            sp = ScheduledPayment(party_type='customer', party_id=customer_id, amount_due=due_amount,
                                  due_date=due_date, direction='in')
            db.session.add(sp)

        db.session.commit()
        flash('Sale recorded', 'success')
        return redirect(url_for('sales'))
    return render_template('billing.html', customers=customers, products=products, now=datetime.now())


@app.route('/suppliers', methods=['GET', 'POST'])
@login_required
def suppliers():
    # handle create or update
    if request.method == 'POST':
        sid = request.form.get('id')
        name = request.form.get('name')
        phone = request.form.get('phone')
        gst = request.form.get('gst')
        if sid:
            s = Supplier.query.get(sid)
            if s:
                s.name = name or s.name
                s.phone = phone or s.phone
                s.gst = gst or s.gst
                db.session.commit()
                flash('Supplier updated', 'success')
                return redirect(url_for('suppliers'))
        s = Supplier(name=name, phone=phone, gst=gst)
        db.session.add(s)
        db.session.commit()
        flash('Supplier added', 'success')
        return redirect(url_for('suppliers'))
    suppliers = Supplier.query.all()
    # if ?id= provided, forward to template to prefill edit
    edit_id = request.args.get('id')
    return render_template('suppliers.html', suppliers=suppliers, edit_id=edit_id)


@app.route('/api/supplier/<int:sid>')
@login_required
def api_supplier(sid):
    s = Supplier.query.get_or_404(sid)
    return {
        'id': s.id,
        'name': s.name,
        'phone': s.phone,
        'gst': s.gst
    }


@app.route('/products', methods=['GET', 'POST'])
@login_required
def products():
    if request.method == 'POST':
        pid = request.form.get('id')
        name = request.form.get('name')
        sku = request.form.get('sku')
        mrp_price = float(request.form.get('mrp_price') or 0)
        quantity = int(request.form.get('quantity') or 0)
        original_quantity = int(request.form.get('original_quantity') or 0)
        purchase_price = float(request.form.get('purchase_price') or 0)
        price = float(request.form.get('price') or 0)
        # Accept CGST/SGST separately or a combined gst_percent for backward compatibility
        try:
            cgst = float(request.form.get('cgst_percent') or request.form.get('cgst') or 0)
        except Exception:
            cgst = 0.0
        try:
            sgst = float(request.form.get('sgst_percent') or request.form.get('sgst') or 0)
        except Exception:
            sgst = 0.0
        gst = cgst + sgst
        # fallback to legacy single gst_percent field if cgst/sgst not provided
        if cgst == 0 and sgst == 0:
            try:
                legacy = float(request.form.get('gst_percent') or 0)
                gst = legacy
                cgst = legacy / 2.0
                sgst = legacy / 2.0
            except Exception:
                pass
        medicine_type = request.form.get('medicine_type') or 'LIQUID'
        tablets_per_strip = int(request.form.get('tablets_per_strip') or 0)
        expiry = request.form.get('expiry_date')
        next_page = request.form.get('next') or request.args.get('next')

        if medicine_type.strip().upper() == 'STRIPS' and tablets_per_strip > 0:
            if pid and original_quantity > 0:
                original_strip_qty = original_quantity // tablets_per_strip
                original_remainder = original_quantity % tablets_per_strip
                if quantity == original_strip_qty:
                    quantity_for_storage = quantity * tablets_per_strip + original_remainder
                else:
                    quantity_for_storage = quantity * tablets_per_strip
            else:
                quantity_for_storage = quantity * tablets_per_strip
        else:
            quantity_for_storage = quantity

        if pid:
            p = Product.query.get(pid)
            if p:
                p.name = name or p.name
                p.sku = sku or p.sku
                p.mrp_price = mrp_price
                p.quantity = quantity_for_storage
                p.purchase_price = purchase_price
                p.price = price
                p.gst_percent = gst
                p.medicine_type = medicine_type
                p.tablets_per_strip = tablets_per_strip
                if expiry:
                    try:
                        p.expiry_date = datetime.strptime(expiry, '%Y-%m-%d').date()
                    except Exception:
                        pass
                db.session.commit()
                flash('Medicine updated', 'success')
                if next_page == 'purchase':
                    return redirect(url_for('purchase'))
                if next_page == 'billing':
                    return redirect(url_for('billing'))
                if next_page and next_page.startswith('/'):
                    return redirect(next_page)
                return redirect(url_for('products'))
        # create new
        p = Product(name=name, sku=sku, mrp_price=mrp_price, quantity=quantity_for_storage, purchase_price=purchase_price, price=price, gst_percent=gst, medicine_type=medicine_type, tablets_per_strip=tablets_per_strip)
        if expiry:
            try:
                # Accept YYYY-MM (month picker) or full YYYY-MM-DD
                if len(expiry) == 7 and '-' in expiry:
                    y, m = expiry.split('-')
                    y = int(y); m = int(m)
                    last_day = calendar.monthrange(y, m)[1]
                    p.expiry_date = date(y, m, last_day)
                else:
                    p.expiry_date = datetime.strptime(expiry, '%Y-%m-%d').date()
            except Exception:
                pass
        db.session.add(p)
        db.session.commit()
        flash('Medicine saved', 'success')
        if next_page == 'purchase':
            return redirect(url_for('purchase'))
        if next_page == 'billing':
            return redirect(url_for('billing'))
        if next_page and next_page.startswith('/'):
            return redirect(next_page)
        return redirect(url_for('products'))
    products = Product.query.all()
    return render_template('products.html', products=products)


@app.route('/delete_product/<int:product_id>', methods=['POST'])
@login_required
def delete_product(product_id):
    product = Product.query.get_or_404(product_id)
    
    # Check if product has any purchases or sales
    has_purchases = Purchase.query.filter_by(product_id=product_id).count() > 0
    has_sales = Sale.query.filter_by(product_id=product_id).count() > 0
    
    if has_purchases or has_sales:
        flash('Cannot delete product: it has associated purchases or sales records', 'danger')
    else:
        try:
            db.session.delete(product)
            db.session.commit()
            flash('Product deleted successfully', 'success')
        except Exception as e:
            db.session.rollback()
            flash(f'Error deleting product: {str(e)}', 'danger')
    
    return redirect(url_for('products'))


@app.route('/sales')
@login_required
def sales():
    sales = Sale.query.order_by(Sale.date.desc()).limit(200).all()
    for s in sales:
        s.customer = Customer.query.get(s.customer_id) if s.customer_id else None
        s.product = Product.query.get(s.product_id) if s.product_id else None
    return render_template('sales.html', sales=sales)


@app.route('/edit_sale/<int:sale_id>', methods=['GET', 'POST'])
@login_required
def edit_sale(sale_id):
    sale = Sale.query.get_or_404(sale_id)
    customers = Customer.query.all()
    products = Product.query.all()
    
    if request.method == 'POST':
        try:
            customer_id = request.form.get('customer')
            product_id = request.form.get('product')
            sku = request.form.get('sku')
            qty = int(request.form.get('qty', 0))
            price = float(request.form.get('price', 0))
            mrp_price = float(request.form.get('mrp_price', 0))
            gst_percent = float(request.form.get('gst_percent', 0))
            discount = float(request.form.get('discount', 0))
            seller_name = request.form.get('seller_name')
            sale_date = request.form.get('date')
            due_date = request.form.get('due_date')
            
            if not customer_id or not product_id or qty <= 0 or price < 0:
                flash('Please fill in all required fields correctly', 'danger')
                return redirect(url_for('edit_sale', sale_id=sale_id))
            
            # Restore old inventory
            old_product = Product.query.get(sale.product_id)
            if old_product:
                old_product.quantity += sale.qty
            
            # Calculate new total
            cgst = gst_percent / 2.0
            sgst = gst_percent / 2.0
            subtotal = qty * price
            gst_amount = subtotal * gst_percent / 100.0
            total = subtotal + gst_amount - discount
            
            # Update sale
            sale.customer_id = int(customer_id)
            sale.product_id = int(product_id)
            sale.sku = sku
            sale.qty = qty
            sale.price = price
            sale.mrp_price = mrp_price
            sale.gst_percent = gst_percent
            sale.cgst_percent = cgst
            sale.sgst_percent = sgst
            sale.discount = discount
            sale.total = total
            sale.seller_name = seller_name
            sale.date = datetime.strptime(sale_date, '%Y-%m-%d').date() if sale_date else date.today()
            sale.due_date = datetime.strptime(due_date, '%Y-%m-%d').date() if due_date else None
            
            # Reduce new product inventory
            new_product = Product.query.get(int(product_id))
            if new_product:
                new_product.quantity = max(0, new_product.quantity - qty)
            
            db.session.commit()
            flash('Sale updated successfully', 'success')
            return redirect(url_for('sales'))
        except Exception as e:
            db.session.rollback()
            flash(f'Error updating sale: {str(e)}', 'danger')
            return redirect(url_for('edit_sale', sale_id=sale_id))
    
    sale.customer = Customer.query.get(sale.customer_id) if sale.customer_id else None
    sale.product = Product.query.get(sale.product_id) if sale.product_id else None
    return render_template('edit_sale.html', sale=sale, customers=customers, products=products)


@app.route('/delete_sale/<int:sale_id>', methods=['POST'])
@login_required
def delete_sale(sale_id):
    sale = Sale.query.get_or_404(sale_id)
    try:
        # Restore inventory
        product = Product.query.get(sale.product_id)
        if product:
            product.quantity += sale.qty
        
        db.session.delete(sale)
        db.session.commit()
        flash('Sale deleted successfully', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error deleting sale: {str(e)}', 'danger')
    
    return redirect(url_for('sales'))


@app.route('/purchase_returns', methods=['GET', 'POST'])
@login_required
def purchase_returns():
    suppliers = Supplier.query.all()
    products = Product.query.all()
    returns = PurchaseReturn.query.order_by(PurchaseReturn.date.desc()).limit(200).all()
    for r in returns:
        r.supplier = Supplier.query.get(r.supplier_id) if r.supplier_id else None
        r.product = Product.query.get(r.product_id) if r.product_id else None

    selected_supplier_id = None
    selected_product_id = None
    default_price = 0
    default_qty = 1
    original_purchase = None

    purchase_id = request.args.get('purchase_id')
    if purchase_id and request.method == 'GET':
        purchase = Purchase.query.get(purchase_id)
        if purchase:
            selected_supplier_id = purchase.supplier_id
            selected_product_id = purchase.product_id
            default_price = purchase.price
            default_qty = purchase.qty
            original_purchase = purchase

    if request.method == 'POST':
        supplier_id = request.form.get('supplier')
        product_id = request.form.get('product')
        qty = int(request.form.get('qty') or 0)
        price = float(request.form.get('price') or 0)
        payment_method = request.form.get('payment_method')
        payment_time = request.form.get('payment_time')
        remark = request.form.get('remark')
        date_val = request.form.get('date')
        date_val = datetime.strptime(date_val, '%Y-%m-%d').date() if date_val else date.today()

        if not supplier_id or not product_id or qty <= 0 or price < 0:
            flash('Please select supplier, product and enter a valid quantity and price.', 'danger')
            return redirect(url_for('purchase_returns'))

        prod = Product.query.get(product_id)
        if not prod:
            flash('Selected product not found.', 'danger')
            return redirect(url_for('purchase_returns'))
        if qty > prod.quantity:
            flash('Return quantity cannot exceed current inventory stock.', 'danger')
            return redirect(url_for('purchase_returns'))

        total = qty * price
        purchase_return = PurchaseReturn(supplier_id=supplier_id, product_id=product_id,
                                         sku=prod.sku, qty=qty, price=price, total=total,
                                         remark=remark, date=date_val)
        prod.quantity = max(0, prod.quantity - qty)
        db.session.add(purchase_return)

        if total > 0:
            payment = Payment(party_type='supplier', party_id=supplier_id, amount=total,
                              mode='refund', payment_method=payment_method,
                              payment_time=datetime.strptime(payment_time, '%H:%M').time() if payment_time else None,
                              direction='in', date=date_val)
            db.session.add(payment)

        db.session.commit()
        check_low_stock_and_alert(prod)
        flash('Purchase return recorded and inventory updated.', 'success')
        return redirect(url_for('purchase_returns'))

    return render_template('purchase_returns.html', suppliers=suppliers, products=products, returns=returns,
                           selected_supplier_id=selected_supplier_id,
                           selected_product_id=selected_product_id,
                           default_price=default_price,
                           default_qty=default_qty,
                           original_purchase=original_purchase)


@app.route('/sale_returns', methods=['GET', 'POST'])
@login_required
def sale_returns():
    customers = Customer.query.all()
    products = Product.query.all()
    returns = SaleReturn.query.order_by(SaleReturn.date.desc()).limit(200).all()
    for r in returns:
        r.customer = Customer.query.get(r.customer_id) if r.customer_id else None
        r.product = Product.query.get(r.product_id) if r.product_id else None

    selected_customer_id = None
    selected_product_id = None
    default_price = 0
    default_qty = 1
    original_sale = None

    sale_id = request.args.get('sale_id')
    if sale_id and request.method == 'GET':
        sale = Sale.query.get(sale_id)
        if sale:
            selected_customer_id = sale.customer_id
            selected_product_id = sale.product_id
            default_price = sale.price
            default_qty = sale.qty
            original_sale = sale

    recent_sales = Sale.query.order_by(Sale.date.desc(), Sale.id.desc()).all()
    customer_prices = {}
    for s in recent_sales:
        key = f"{s.customer_id}_{s.product_id}"
        if key not in customer_prices:
            customer_prices[key] = s.price

    if request.method == 'POST':
        customer_id = request.form.get('customer')
        product_id = request.form.get('product')
        qty = int(request.form.get('qty') or 0)
        price = float(request.form.get('price') or 0)
        payment_method = request.form.get('payment_method')
        payment_time = request.form.get('payment_time')
        remark = request.form.get('remark')
        date_val = request.form.get('date')
        date_val = datetime.strptime(date_val, '%Y-%m-%d').date() if date_val else date.today()

        if not customer_id or not product_id or qty <= 0 or price < 0:
            flash('Please select customer, product and enter a valid quantity and price.', 'danger')
            return redirect(url_for('sale_returns'))

        prod = Product.query.get(product_id)
        if not prod:
            flash('Selected product not found.', 'danger')
            return redirect(url_for('sale_returns'))

        total = qty * price
        sale_return = SaleReturn(customer_id=customer_id, product_id=product_id,
                                 sku=prod.sku, qty=qty, price=price, total=total,
                                 remark=remark, date=date_val)
        prod.quantity += qty
        db.session.add(sale_return)

        if total > 0:
            payment = Payment(party_type='customer', party_id=customer_id, amount=total,
                              mode='refund', payment_method=payment_method,
                              payment_time=datetime.strptime(payment_time, '%H:%M').time() if payment_time else None,
                              direction='out', date=date_val)
            db.session.add(payment)

        db.session.commit()
        flash('Sale return recorded, refund logged and inventory updated.', 'success')
        return redirect(url_for('sale_returns'))

    return render_template('sale_returns.html', customers=customers, products=products, returns=returns,
                           customer_prices=customer_prices,
                           selected_customer_id=selected_customer_id,
                           selected_product_id=selected_product_id,
                           default_price=default_price,
                           default_qty=default_qty,
                           original_sale=original_sale)


@app.route('/payments', methods=['GET', 'POST'])
@login_required
def payments():
    if request.method == 'POST':
        party_type = request.form.get('party_type')
        party_id = request.form.get('party_id')
        amount = float(request.form.get('amount') or 0)
        mode = request.form.get('mode')
        payment_method = request.form.get('payment_method')
        payment_time_value = request.form.get('payment_time')
        direction = request.form.get('direction')
        date_val = request.form.get('date')
        payment = Payment(party_type=party_type, party_id=party_id, amount=amount, mode=mode,
                          payment_method=payment_method,
                          payment_time=datetime.strptime(payment_time_value, '%H:%M').time() if payment_time_value else None,
                          direction=direction,
                          date=datetime.strptime(date_val, '%Y-%m-%d').date() if date_val else date.today())
        db.session.add(payment)
        # if this pays a scheduled payment, mark it paid
        sp_id = request.form.get('scheduled_payment_id')
        if sp_id:
            sp = ScheduledPayment.query.get(sp_id)
            if sp:
                sp.status = 'paid'
        db.session.commit()
        flash('Payment recorded', 'success')
        return redirect(url_for('payments'))

    customers = Customer.query.all()
    suppliers = Supplier.query.all()
    scheduled = ScheduledPayment.query.filter_by(status='pending').all()
    payments = Payment.query.order_by(Payment.date.desc()).limit(50).all()
    return render_template('payments.html', customers=customers, suppliers=suppliers, scheduled=scheduled, payments=payments)


@app.route('/reminders')
@login_required
def reminders():
    reminders = Reminder.query.order_by(Reminder.date.desc()).all()
    pending = ScheduledPayment.query.filter(ScheduledPayment.status=='pending').all()
    return render_template('reminders.html', reminders=reminders, pending=pending)


@app.route('/notifications')
@login_required
def notifications():
    uid = session['user_id']
    notifs = Notification.query.filter_by(user_id=uid).order_by(Notification.created.desc()).all()
    unread_count = Notification.query.filter_by(user_id=uid, read=False).count()
    return render_template('notifications.html', notifications=notifs, unread_count=unread_count)


@app.route('/notification/<int:notif_id>/read', methods=['POST'])
@login_required
def mark_notification_read(notif_id):
    notif = Notification.query.get_or_404(notif_id)
    if notif.user_id == session['user_id']:
        notif.read = True
        db.session.commit()
    return {'status': 'ok'}


@app.route('/notifications/unread-count')
@login_required
def unread_count():
    uid = session['user_id']
    count = Notification.query.filter_by(user_id=uid, read=False).count()
    return {'count': count}


def check_scheduled_payments():
    with app.app_context():
        # create reminders for scheduled payments whose due_date is today or earlier and still pending
        today = date.today()
        due = ScheduledPayment.query.filter(ScheduledPayment.status=='pending').filter(ScheduledPayment.due_date != None).filter(ScheduledPayment.due_date <= today).all()
        for sp in due:
            existing = Reminder.query.filter_by(scheduled_payment_id=sp.id, date=today).first()
            if not existing:
                msg = f"Payment due: {sp.party_type} {sp.party_id} amount ₹{sp.amount_due} due {sp.due_date}"
                r = Reminder(scheduled_payment_id=sp.id, message=msg, date=today, sent=False)
                db.session.add(r)
        db.session.commit()
        # try to send reminders via email if configured
        for r in Reminder.query.filter_by(date=today, sent=False).all():
            try:
                send_email_subject = 'Payment Reminder'
                send_email_body = r.message
                send_email(send_email_subject, send_email_body)
                r.sent = True
                db.session.commit()
            except Exception:
                # skip failures; they'll be retried next run
                pass


def send_email(subject: str, body: str):
    host = os.environ.get('SMTP_HOST')
    port = int(os.environ.get('SMTP_PORT', '587'))
    user = os.environ.get('SMTP_USER')
    pwd = os.environ.get('SMTP_PASS')
    to_addr = os.environ.get('ALERT_EMAIL_TO')
    if not host or not to_addr:
        raise RuntimeError('SMTP not configured')
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = user or 'noreply@example.com'
    msg['To'] = to_addr
    msg.set_content(body)
    with smtplib.SMTP(host, port) as s:
        s.starttls()
        if user and pwd:
            s.login(user, pwd)
        s.send_message(msg)


def check_low_stock_and_alert(product: Product, threshold: int = 5):
    if product.quantity <= threshold:
        # create a reminder entry for low stock
        msg = f"Low stock alert: {product.name} (id {product.id}) qty {product.quantity}"
        r = Reminder(scheduled_payment_id=None, message=msg, date=date.today(), sent=False)
        db.session.add(r)
        db.session.commit()
        # try sending email
        try:
            send_email('Low Stock Alert', msg)
            r.sent = True
            db.session.commit()
        except Exception:
            pass
        # create notification for all admin users
        admins = User.query.filter_by(role='admin').all()
        for admin in admins:
            notif = Notification(user_id=admin.id, title='Low Stock Alert', message=msg, 
                                notification_type='low_stock', related_id=product.id, 
                                action_url=f'/inventory')
            db.session.add(notif)
        db.session.commit()


def create_notification(user_id: int, title: str, message: str, notif_type: str = 'info', 
                       related_id: int = None, action_url: str = None):
    notif = Notification(user_id=user_id, title=title, message=message, notification_type=notif_type,
                        related_id=related_id, action_url=action_url)
    db.session.add(notif)
    db.session.commit()
    return notif


@app.route('/shop-profile', methods=['GET', 'POST'])
@role_required(['admin'])
def shop_profile():
    shop = ShopProfile.query.first()
    if not shop:
        shop = create_default_shop_profile()
    if request.method == 'POST':
        shop.name = request.form.get('name') or shop.name
        shop.gst = request.form.get('gst')
        shop.address = request.form.get('address')
        shop.phone = request.form.get('phone')
        shop.email = request.form.get('email')
        db.session.commit()
        flash('Shop profile updated', 'success')
        return redirect(url_for('shop_profile'))
    return render_template('shop_profile.html', shop=shop)


@app.route('/customers', methods=['GET', 'POST'])
@login_required
def customers():
    if request.method == 'POST':
        cid = request.form.get('id')
        name = request.form.get('name')
        phone = request.form.get('phone')
        address = request.form.get('address')
        payment_due = float(request.form.get('payment_due') or 0)
        payment_complete = request.form.get('payment_complete') == 'on'
        remark = request.form.get('remark')
        if cid:
            c = Customer.query.get(cid)
            if c:
                c.name = name or c.name
                c.phone = phone or c.phone
                c.address = address or c.address
                c.payment_due = payment_due
                c.payment_complete = payment_complete
                c.remark = remark or c.remark
                db.session.commit()
                flash('Customer updated', 'success')
                return redirect(url_for('customers'))
        c = Customer(name=name, phone=phone, address=address, payment_due=payment_due,
                     payment_complete=payment_complete, remark=remark)
        db.session.add(c)
        db.session.commit()
        flash('Customer added', 'success')
        return redirect(url_for('customers'))
    customers = Customer.query.all()
    edit_id = request.args.get('id')
    return render_template('customers.html', customers=customers, edit_id=edit_id)


@app.route('/api/customer/<int:cid>')
@login_required
def api_customer(cid):
    c = Customer.query.get_or_404(cid)
    return {
        'id': c.id,
        'name': c.name,
        'phone': c.phone,
        'address': c.address,
        'payment_due': c.payment_due,
        'payment_complete': c.payment_complete,
        'remark': c.remark,
    }


@app.route('/reports', methods=['GET'])
@login_required
def reports():
    from_date = request.args.get('from')
    to_date = request.args.get('to')
    sales_q = Sale.query
    purchases_q = Purchase.query
    purchase_returns_q = PurchaseReturn.query
    sale_returns_q = SaleReturn.query

    if from_date:
        d1 = datetime.strptime(from_date, '%Y-%m-%d').date()
        sales_q = sales_q.filter(Sale.date >= d1)
        purchases_q = purchases_q.filter(Purchase.date >= d1)
        purchase_returns_q = purchase_returns_q.filter(PurchaseReturn.date >= d1)
        sale_returns_q = sale_returns_q.filter(SaleReturn.date >= d1)
    if to_date:
        d2 = datetime.strptime(to_date, '%Y-%m-%d').date()
        sales_q = sales_q.filter(Sale.date <= d2)
        purchases_q = purchases_q.filter(Purchase.date <= d2)
        purchase_returns_q = purchase_returns_q.filter(PurchaseReturn.date <= d2)
        sale_returns_q = sale_returns_q.filter(SaleReturn.date <= d2)

    sales = sales_q.all()
    purchases = purchases_q.all()
    purchase_returns = purchase_returns_q.order_by(PurchaseReturn.date.desc()).all()
    sale_returns = sale_returns_q.order_by(SaleReturn.date.desc()).all()

    for r in purchase_returns:
        r.supplier = Supplier.query.get(r.supplier_id) if r.supplier_id else None
        r.product = Product.query.get(r.product_id) if r.product_id else None
    for r in sale_returns:
        r.customer = Customer.query.get(r.customer_id) if r.customer_id else None
        r.product = Product.query.get(r.product_id) if r.product_id else None

    total_sales = sum(s.total for s in sales)
    total_purchases = sum(p.total for p in purchases)
    total_purchase_returns = sum(r.total for r in purchase_returns)
    total_sale_returns = sum(r.total for r in sale_returns)
    profit = total_sales - total_purchases - total_sale_returns + total_purchase_returns
    return render_template('reports.html', sales=sales, purchases=purchases,
                           purchase_returns=purchase_returns, sale_returns=sale_returns,
                           total_purchase_returns=total_purchase_returns,
                           total_sale_returns=total_sale_returns,
                           profit=profit)


@app.route('/export', methods=['GET'])
@login_required
def export_page():
    """Render Export profile page where user selects dataset/year/month."""
    years = set()
    try:
        for model in (Sale, Purchase, Payment, PurchaseReturn, SaleReturn):
            rows = db.session.query(model.date).filter(model.date != None).all()
            for (d,) in rows:
                try:
                    years.add(d.year)
                except Exception:
                    continue
    except Exception:
        years = set()

    years_list = sorted(years, reverse=True)
    return render_template('export.html', years=years_list)


@app.route('/print/<int:sale_id>')
@app.route('/print_bill/<int:sale_id>')
@login_required
def print_bill(sale_id):
    sale = Sale.query.get_or_404(sale_id)
    product = Product.query.get(sale.product_id)
    customer = Customer.query.get(sale.customer_id)
    return render_template('print_bill.html', sale=sale, product=product, customer=customer)


@app.route('/print/purchase_return/<int:return_id>')
@login_required
def print_purchase_return(return_id):
    purchase_return = PurchaseReturn.query.get_or_404(return_id)
    product = Product.query.get(purchase_return.product_id)
    supplier = Supplier.query.get(purchase_return.supplier_id)
    return render_template('print_return.html', return_record=purchase_return,
                           entity=supplier, product=product,
                           return_type='purchase')


@app.route('/print/sale_return/<int:return_id>')
@login_required
def print_sale_return(return_id):
    sale_return = SaleReturn.query.get_or_404(return_id)
    product = Product.query.get(sale_return.product_id)
    customer = Customer.query.get(sale_return.customer_id)
    return render_template('print_return.html', return_record=sale_return,
                           entity=customer, product=product,
                           return_type='sale')


# ==================== EXCEL EXPORT ROUTES ====================

@app.route('/export/purchases')
@login_required
def export_purchases_route():
    """Export purchases data to Excel with year & month wise organization"""
    from_date = request.args.get('from')
    to_date = request.args.get('to')
    
    query = Purchase.query
    
    if from_date:
        try:
            d1 = datetime.strptime(from_date, '%Y-%m-%d').date()
            query = query.filter(Purchase.date >= d1)
        except:
            pass
    
    if to_date:
        try:
            d2 = datetime.strptime(to_date, '%Y-%m-%d').date()
            query = query.filter(Purchase.date <= d2)
        except:
            pass
    
    purchases = query.order_by(Purchase.date.desc()).all()
    
    # Eager load relationships
    for purchase in purchases:
        purchase.supplier = Supplier.query.get(purchase.supplier_id)
        purchase.product = Product.query.get(purchase.product_id)
    
    wb = export_purchases(purchases)
    file_bytes = excel_to_bytes(wb)
    
    filename = f"Purchases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/sales')
@login_required
def export_sales_route():
    """Export sales data to Excel with year & month wise organization"""
    from_date = request.args.get('from')
    to_date = request.args.get('to')
    
    query = Sale.query
    
    if from_date:
        try:
            d1 = datetime.strptime(from_date, '%Y-%m-%d').date()
            query = query.filter(Sale.date >= d1)
        except:
            pass
    
    if to_date:
        try:
            d2 = datetime.strptime(to_date, '%Y-%m-%d').date()
            query = query.filter(Sale.date <= d2)
        except:
            pass
    
    sales = query.order_by(Sale.date.desc()).all()
    
    # Eager load relationships
    for sale in sales:
        sale.customer = Customer.query.get(sale.customer_id)
        sale.product = Product.query.get(sale.product_id)
    
    wb = export_sales(sales)
    file_bytes = excel_to_bytes(wb)
    
    filename = f"Sales_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument-spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/purchase_returns')
@login_required
def export_purchase_returns_route():
    from_date = request.args.get('from')
    to_date = request.args.get('to')
    query = PurchaseReturn.query

    if from_date:
        try:
            d1 = datetime.strptime(from_date, '%Y-%m-%d').date()
            query = query.filter(PurchaseReturn.date >= d1)
        except:
            pass

    if to_date:
        try:
            d2 = datetime.strptime(to_date, '%Y-%m-%d').date()
            query = query.filter(PurchaseReturn.date <= d2)
        except:
            pass

    returns = query.order_by(PurchaseReturn.date.desc()).all()
    for r in returns:
        r.supplier = Supplier.query.get(r.supplier_id)
        r.product = Product.query.get(r.product_id)

    wb = export_purchase_returns(returns)
    file_bytes = excel_to_bytes(wb)
    filename = f"PurchaseReturns_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/sale_returns')
@login_required
def export_sale_returns_route():
    from_date = request.args.get('from')
    to_date = request.args.get('to')
    query = SaleReturn.query

    if from_date:
        try:
            d1 = datetime.strptime(from_date, '%Y-%m-%d').date()
            query = query.filter(SaleReturn.date >= d1)
        except:
            pass

    if to_date:
        try:
            d2 = datetime.strptime(to_date, '%Y-%m-%d').date()
            query = query.filter(SaleReturn.date <= d2)
        except:
            pass

    returns = query.order_by(SaleReturn.date.desc()).all()
    for r in returns:
        r.customer = Customer.query.get(r.customer_id)
        r.product = Product.query.get(r.product_id)

    wb = export_sale_returns(returns)
    file_bytes = excel_to_bytes(wb)
    filename = f"SaleReturns_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument-spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/products')
@login_required
def export_products_route():
    """Export products/inventory data to Excel"""
    products = Product.query.all()
    wb = export_products(products)
    file_bytes = excel_to_bytes(wb)
    
    filename = f"Inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/suppliers')
@login_required
def export_suppliers_route():
    """Export suppliers data to Excel"""
    suppliers = Supplier.query.all()
    wb = export_suppliers(suppliers)
    file_bytes = excel_to_bytes(wb)
    
    filename = f"Suppliers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/customers')
@login_required
def export_customers_route():
    """Export customers data to Excel"""
    customers = Customer.query.all()
    wb = export_customers(customers)
    file_bytes = excel_to_bytes(wb)
    
    filename = f"Customers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/payments')
@login_required
def export_payments_route():
    """Export payments data to Excel with year & month wise organization"""
    from_date = request.args.get('from')
    to_date = request.args.get('to')
    
    query = Payment.query
    
    if from_date:
        try:
            d1 = datetime.strptime(from_date, '%Y-%m-%d').date()
            query = query.filter(Payment.date >= d1)
        except:
            pass
    
    if to_date:
        try:
            d2 = datetime.strptime(to_date, '%Y-%m-%d').date()
            query = query.filter(Payment.date <= d2)
        except:
            pass
    
    payments = query.order_by(Payment.date.desc()).all()
    wb = export_payments(payments)
    file_bytes = excel_to_bytes(wb)
    
    filename = f"Payments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/export/bills')
@login_required
def export_bills_route():
    """Export bills data to Excel with year & month wise organization"""
    from_date = request.args.get('from')
    to_date = request.args.get('to')
    
    query = Sale.query
    
    if from_date:
        try:
            d1 = datetime.strptime(from_date, '%Y-%m-%d').date()
            query = query.filter(Sale.date >= d1)
        except:
            pass
    
    if to_date:
        try:
            d2 = datetime.strptime(to_date, '%Y-%m-%d').date()
            query = query.filter(Sale.date <= d2)
        except:
            pass
    
    bills = query.order_by(Sale.date.desc()).all()
    
    # Eager load relationships
    for bill in bills:
        bill.customer = Customer.query.get(bill.customer_id)
        bill.product = Product.query.get(bill.product_id)
    
    wb = export_bills(bills)
    file_bytes = excel_to_bytes(wb)
    
    filename = f"Bills_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


# Run initialization after all app routes and helpers are defined.
# Disable scheduler when running under pytest.
with app.app_context():
    running_under_test = 'PYTEST_CURRENT_TEST' in os.environ or 'pytest' in sys.modules
    init_db(start_scheduler=not running_under_test)


if __name__ == '__main__':
    debug_mode = os.environ.get('FLASK_DEBUG', 'false').lower() in ('1', 'true', 'yes', 'on')
    app.run(debug=debug_mode, use_reloader=False)
