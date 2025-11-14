import sqlite3
import random
from datetime import datetime
from flask import Flask, request, session, redirect, url_for, g

# --- Konfigurasi Aplikasi ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'kunci-rahasia-yang-sangat-aman-12345'
DATABASE = 'accounting.db'

# --- Fungsi Database (SQLite3) ---

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def generate_journal_code(db, entry_datetime=None):
    """
    Membuat ID Jurnal unik dengan format: SG-YYYYMMDD-HHMMSS-RANDOM5
    entry_datetime: objek datetime (jika manual) atau None (untuk realtime)
    """
    if entry_datetime is None:
        entry_datetime = datetime.now()
    
    date_str = entry_datetime.strftime('%Y%m%d')
    time_str = entry_datetime.strftime('%H%M%S')
    
    # Loop untuk menjamin keunikan (meski kemungkinan tabrakan sangat kecil)
    while True:
        rand_str = f"{random.randint(10000, 99999)}"
        new_code = f"SG{date_str}{time_str}{rand_str}"
        
        # Cek ke database apakah kode ini sudah ada
        exists = db.execute("SELECT 1 FROM journal_entries WHERE journal_code = ?", (new_code,)).fetchone()
        if not exists:
            return new_code

def init_db():
    print("Membuat database...")
    with app.app_context():
        db = get_db()
        db.execute("PRAGMA foreign_keys = ON;")
        
        with db:
            # ... (Tabel users, chart_of_accounts, journal_entries, journal_details dari kode sebelumnya) ...
            db.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin', 'consumer')),
                security_answer TEXT NOT NULL
            )
            ''')
            db.execute('''
            CREATE TABLE IF NOT EXISTS chart_of_accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_code TEXT UNIQUE NOT NULL,
                account_name TEXT NOT NULL,
                account_type TEXT NOT NULL CHECK(account_type IN ('Aset Tetap', 'Aset Lancar', 'Liabilitas', 'Ekuitas', 'Pendapatan', 'Beban'))
            )
            ''')
            db.execute('''
            CREATE TABLE IF NOT EXISTS journal_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                journal_code TEXT UNIQUE NOT NULL,
                entry_timestamp TEXT DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')),
                description TEXT NOT NULL
            )
            ''')
            db.execute('''
            CREATE TABLE IF NOT EXISTS journal_details (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_id INTEGER NOT NULL,
                account_code TEXT NOT NULL,
                debit REAL DEFAULT 0,
                credit REAL DEFAULT 0,
                FOREIGN KEY (entry_id) REFERENCES journal_entries (id) ON DELETE CASCADE,
                FOREIGN KEY (account_code) REFERENCES chart_of_accounts (account_code) ON DELETE RESTRICT
            )
            ''')

            # --- TABEL BARU UNTUK INVENTORY ---
            db.execute('''
            CREATE TABLE IF NOT EXISTS inventory_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_name TEXT UNIQUE NOT NULL,
                # Akun Persediaan & HPP yang terkait
                inventory_account TEXT NOT NULL, 
                cogs_account TEXT NOT NULL,
                sales_account TEXT NOT NULL,
                FOREIGN KEY (inventory_account) REFERENCES chart_of_accounts (account_code),
                FOREIGN KEY (cogs_account) REFERENCES chart_of_accounts (account_code),
                FOREIGN KEY (sales_account) REFERENCES chart_of_accounts (account_code)
            )
            ''')
            
            db.execute('''
            CREATE TABLE IF NOT EXISTS inventory_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                trx_date DATE NOT NULL,
                trx_type TEXT NOT NULL CHECK(trx_type IN ('purchase', 'sale')),
                description TEXT,
                qty REAL NOT NULL,
                # 'cost' hanya diisi saat 'purchase'
                cost_per_unit REAL DEFAULT 0,
                # 'price' hanya diisi saat 'sale' (untuk jurnal pendapatan)
                sale_price_per_unit REAL DEFAULT 0,
                FOREIGN KEY (item_id) REFERENCES inventory_items (id)
            )
            ''')
            # --- AKHIR TABEL BARU ---

            try:
                # ... (Data users dan chart_of_accounts) ...
                db.execute("INSERT INTO users (username, password, role, security_answer) VALUES (?, ?, ?, ?)", 
                           ('admin', 'admin123', 'admin', 'admin@mail.com'))
                db.execute("INSERT INTO users (username, password, role, security_answer) VALUES (?, ?, ?, ?)", 
                           ('consumer', 'consumer123', 'consumer', 'consumer@mail.com'))
                
                # (Daftar akun Anda akan di-insert di sini seperti biasa)
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1101', 'Kas', 'Aset Lancar'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1102', 'Piutang Dagang', 'Aset Lancar'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1103', 'Perlengkapan', 'Aset Lancar'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1104', 'Peralatan', 'Aset Lancar'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1105', 'Persediaan', 'Aset Lancar'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1105', 'Persediaan Pakan', 'Aset Lancar'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1201', 'Peralatan', 'Aset Tetap'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1202', 'Akumulasi Penyusutan Peralatan', 'Aset Tetap'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1203', 'Kendaraan', 'Aset Tetap'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1204', 'Akumulasi Penyusutan Kendaraan', 'Aset Tetap'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1205', 'Bangunan', 'Aset Tetap'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1206', 'Akumulasi Penyusutan Bangunan', 'Aset Tetap'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('1207', 'Tanah', 'Aset Tetap'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('2101', 'Hutang Dagang', 'Liabilitas'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('3101', 'Modal', 'Ekuitas'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('3102', 'Prive', 'Ekuitas'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('4101', 'Penjualan', 'Pendapatan'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('4102', 'Pendapatan Lainnya', 'Pendapatan'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('5101', 'Beban Gaji', 'Beban'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('5102', 'Beban Akomodasi', 'Beban'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('5103', 'Beban Listrik dan Air', 'Beban'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('5104', 'Biaya Penyusutan Peralatan', 'Beban'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('5105', 'Biaya Penyusutan Kendaraan', 'Beban'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('5106', 'Biaya Penyusutan Bangunan', 'Beban'))
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)", ('5107', 'Beban Perlengkapan', 'Beban'))

                # --- AKHIR DATA BARU ---

                print("Data awal berhasil dibuat.")
            except sqlite3.IntegrityError as e:
                print(f"Data awal sudah ada atau error: {e}")
    print("Inisialisasi database selesai.")

# --- Fungsi Bantuan Render Halaman (Pure Python ke HTML) ---

def render_page(title, body_content, sidebar_content=None, error_message=None):
    """
    Merender halaman HTML lengkap sebagai string Python.
    SEKARANG MENANGANI LOGIKA SIDEBAR BUKA/TUTUP.
    """
    
    # Periksa query parameter ?sidebar=...
    # Jika 'closed', maka sidebar_visible = False
    sidebar_visible = request.args.get('sidebar', 'open') != 'closed'
    
    sidebar_html = ""
    main_content_style = "width: 100%;"
    open_close_link = "" # Tautan untuk Buka/Tutup
    error_html = ""

    if error_message:
        # Tampilkan pesan error jika ada
        error_html = f'<p style="color: red; border: 1px solid red; padding: 10px;"><b>ERROR:</b> {error_message}</p>'

    # Logika untuk menampilkan sidebar atau tautan "Buka Sidebar"
    if sidebar_content:
        current_path = request.path # Dapatkan URL saat ini (misal: /admin/general-journal)
        
        if sidebar_visible:
            # TAMPILKAN SIDEBAR
            main_content_style = "width: 75%; float: right; padding-left: 20px;"
            # Buat tautan "Tutup" yang mengarah ke URL saat ini + ?sidebar=closed
            close_link = f'<p style="margin-top: 20px;"><a href="{current_path}?sidebar=closed">Tutup Sidebar</a></p>'
            
            sidebar_html = f"""
            <div style="width: 20%; float: left; border-right: 1px solid #ccc; padding: 10px; height: 100vh;">
                {sidebar_content}
                {close_link}
            </div>
            """
        else:
            # JANGAN TAMPILKAN SIDEBAR
            # Buat tautan "Buka" yang mengarah ke URL saat ini (tanpa query parameter)
            open_close_link = f'<p><a href="{current_path}">Buka Sidebar</a></p>'

    html = f"""
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-M">
        <title>{title} - Perusahaan XYZ</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 0; padding: 0; }}
            h1, h2, h3 {{ color: #333; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .container {{ width: 90%; margin: 20px auto; }}
            .navbar {{ background-color: #333; overflow: hidden; padding: 10px; }}
            .navbar a {{ float: left; color: white; text-decoration: none; padding: 14px 20px; }}
            .navbar a.right {{ float: right; }}
            form {{ border: 1px solid #ccc; padding: 20px; border-radius: 5px; background-color: #f9f9f9; }}
            
            /* Ganti 'width: 90%' menjadi 'width: calc(100% - 22px)' agar pas */
            input[type=text], input[type=password], input[type=number], input[type=date], select {{
                width: calc(100% - 22px); padding: 10px; margin: 10px 0; border: 1px solid #ccc; border-radius: 4px;
            }}
            
            /* Style untuk tombol */
            input[type=submit], input[type=button] {{
                background-color: #4CAF50; color: white; padding: 14px 20px;
                margin: 8px 0; border: none; border-radius: 4px; cursor: pointer;
            }}
            input[type=submit]:hover {{ background-color: #45a049; }}
            
            /* Style khusus untuk tombol Aksi (Biru) dan Hapus (Merah) */
            .btn-blue {{ background-color: #007BFF; }}
            .btn-blue:hover {{ background-color: #0069D9; }}
            .btn-red {{ background-color: #DC3545; }}
            .btn-red:hover {{ background-color: #C82333; }}
            
            /* --- TAMBAHKAN KODE INI --- */
            /* Style untuk kolom sempit */
            .table-condensed th.col-pilih,
            .table-condensed td.col-pilih {{
                width: 5%;
                text-align: center;
            }}
            .table-condensed th.col-nomor,
            .table-condensed td.col-nomor {{
                width: 10%;
                text-align: center;
            }}
            /* --- AKHIR KODE BARU --- */
        </style>
    </head>
    <body>
        <div class="navbar">
            <a href="/"><b>Sistem Akuntansi Perusahaan XYZ</b></a>
            {'<a href="/logout" class="right">Logout</a>' if 'username' in session else ''}
        </div>

        <div class="container">
            {sidebar_html}
            <div style="{main_content_style}">
                {open_close_link}
                {error_html}
                <h1>{title}</h1>
                {body_content}
            </div>
            <div style="clear: both;"></div>
        </div>
    </body>
    </html>
    """
    return html

@app.route("/")
def index():
    if 'username' in session:
        if session.get('role') == 'admin':
            return redirect(url_for('admin_dashboard'))
        elif session.get('role') == 'consumer':
            return redirect(url_for('consumer_home'))
            
    # Cek pesan sukses atau error dari query URL
    success_message = request.args.get('success')
    error_message = request.args.get('error')
    
    message_html = ""
    if success_message:
        message_html = f'<p style="color: green; border: 1px solid green; padding: 10px;">{success_message}</p>'
    if error_message:
        message_html = f'<p style="color: red; border: 1px solid red; padding: 10px;"><b>ERROR:</b> {error_message}</p>'

    body = f"""
    {message_html}
    <h2>Silakan Login</h2>
    
    <form action="/login" method="POST">
        <label for="username">Username:</label>
        <input type="text" id="username" name="username" required autocomplete="username">
        
        <label for="password">Password:</label>
        <input type="password" id="password" name="password" required autocomplete="current-password">
        
        <div style="margin-top: 10px;">
            <a href="/forgot-password">Lupa Sandi?</a>
        </div>
        
        <input type="submit" value="Login">
    </form>
    
    <hr style="margin-top: 20px;">
    <p>Belum punya akun? <a href="/register">Daftar di sini</a></p>
    
    <p><i>Hint Admin: 'admin'/'admin123' (Jawaban: 'admin@mail.com')</i></p>
    """
    return render_page("Login", body)

@app.route("/forgot-password", methods=['GET', 'POST'])
def forgot_password():
    """Halaman untuk verifikasi keamanan (jawaban rahasia/email)."""
    
    error = None
    if request.method == 'POST':
        # 1. Ambil data dari form
        username = request.form['username']
        security_answer = request.form['security_answer']
        
        db = get_db()
        # 2. Cek ke database
        user = db.execute("SELECT * FROM users WHERE username = ? AND security_answer = ?", 
                          (username, security_answer)).fetchone()
        
        if user:
            # 3. JIKA BERHASIL: Simpan di session & redirect ke halaman reset
            session['user_to_reset'] = user['username']
            return redirect(url_for('reset_password'))
        else:
            # 4. JIKA GAGAL: Tampilkan error
            error = "Username atau Jawaban Keamanan salah."

    # Tampilkan form (GET request atau jika POST gagal)
    body = f"""
    <h2>Lupa Sandi</h2>
    <p>Silakan masukkan username dan jawaban keamanan (email) Anda yang terdaftar.</p>
    
    <form action="/forgot-password" method="POST">
        <label for="username">Username:</label>
        <input type="text" id="username" name="username" required>
        
        <label for="security_answer">Jawaban Keamanan (Email):</label>
        <input type="text" id="security_answer" name="security_answer" required>
        
        <input type="submit" value="Verifikasi">
    </form>
    <a href="{url_for('index')}">&larr; Kembali ke Halaman Login</a>
    """
    return render_page("Lupa Sandi", body, error_message=error)

@app.route("/reset-password", methods=['GET', 'POST'])
def reset_password():
    """Halaman untuk memasukkan password baru."""
    
    # 1. Pastikan pengguna sudah lolos verifikasi (dari session)
    if 'user_to_reset' not in session:
        return redirect(url_for('index', error="Sesi reset telah habis. Silakan ulangi."))
    
    username = session['user_to_reset']
    error = None
    
    if request.method == 'POST':
        # 2. Ambil password baru
        new_pass = request.form['new_password']
        confirm_pass = request.form['confirm_password']
        
        # 3. Validasi
        if new_pass != confirm_pass:
            error = "Password baru tidak cocok."
        elif not new_pass:
            error = "Password tidak boleh kosong."
        else:
            # 4. Update database
            db = get_db()
            with db:
                db.execute("UPDATE users SET password = ? WHERE username = ?", (new_pass, username))
            
            # 5. Hapus session & redirect ke login
            session.pop('user_to_reset', None)
            return redirect(url_for('index', success="Password berhasil direset! Silakan login."))

    # Tampilkan form (GET request atau jika POST gagal)
    body = f"""
    <h2>Reset Password untuk: {username}</h2>
    <p>Anda telah terverifikasi. Silakan masukkan password baru Anda.</p>
    
    <form action="/reset-password" method="POST">
        <label for="new_password">Password Baru:</label>
        <input type="password" id="new_password" name="new_password" required>
        
        <label for="confirm_password">Konfirmasi Password Baru:</label>
        <input type="password" id="confirm_password" name="confirm_password" required>
        
        <input type="submit" value="Simpan Password Baru">
    </form>
    """
    return render_page("Reset Password", body, error_message=error)

    # Tampilkan form (GET request atau jika POST gagal)
    body = f"""
    <h2>Reset Password untuk: {username}</h2>
    <p>Anda telah terverifikasi. Silakan masukkan password baru Anda.</p>
    
    <form action="/reset-password" method="POST">
        <label for="new_password">Password Baru:</label>
        <input type="password" id="new_password" name="new_password" required>
        
        <label for="confirm_password">Konfirmasi Password Baru:</label>
        <input type="password" id="confirm_password" name="confirm_password" required>
        
        <input type="submit" value="Simpan Password Baru">
    </form>
    """
    return render_page("Reset Password", body, error_message=error)

@app.route("/login", methods=['POST'])
def login():
    username = request.form['username']
    password = request.form['password']
    
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    
    if user and user['password'] == password:
        session['username'] = user['username']
        session['role'] = user['role']
        if user['role'] == 'admin':
            return redirect(url_for('admin_dashboard'))
        else:
            return redirect(url_for('consumer_home'))
    else:
        # --- PERUBAHAN DI SINI ---
        # Jika data tidak ada atau password salah, kirim pesan error
        return redirect(url_for('index', error="Username atau Password salah."))
        # --- PERUBAHAN SELESAI ---

@app.route("/register", methods=['GET', 'POST'])
def register():
    """Menangani halaman registrasi pengguna baru."""
    
    if request.method == 'POST':
        # 1. Ambil data dari form
        username = request.form['username']
        password = request.form['password']
        confirm_password = request.form['confirm_password']
        security_answer = request.form['security_answer'] # Ini "email"
        
        # 2. Validasi
        if password != confirm_password:
            return render_register_form(error="Password tidak cocok.")
        
        if not username or not password or not security_answer:
            return render_register_form(error="Semua kolom wajib diisi.")
            
        # 3. Simpan ke database
        db = get_db()
        try:
            with db:
                # Otomatis daftarkan pengguna baru sebagai 'consumer'
                db.execute(
                    "INSERT INTO users (username, password, role, security_answer) VALUES (?, ?, ?, ?)",
                    (username, password, 'consumer', security_answer)
                )
            # 4. Jika sukses, kembali ke login
            return redirect(url_for('index', success="Registrasi berhasil! Silakan login."))
            
        except sqlite3.IntegrityError:
            # Ini akan error jika username sudah ada (karena UNIQUE)
            return render_register_form(error="Username ini sudah terdaftar. Silakan pilih nama lain.")

    # Jika method GET, tampilkan form
    return render_register_form()

def render_register_form(error=None):
    """Helper untuk merender form registrasi."""
    
    body = f"""
    <h2>Registrasi Akun Baru</h2>
    <p>Daftar sebagai konsumen baru.</p>
    
    <form action="/register" method="POST">
        <label for="username">Username Baru:</label>
        <input type="text" id="username" name="username" required autocomplete="username">
        
        <label for="password">Password:</label>
        <input type="password" id="password" name="password" required autocomplete="new-password">
        
        <label for="confirm_password">Konfirmasi Password:</label>
        <input type="password" id="confirm_password" name="confirm_password" required autocomplete="new-password">
        
        <label for="security_answer">Email Anda:</label>
        <p style="font-size: 0.9em; color: #555;">Ini akan digunakan jika Anda lupa password. (Contoh: 'emailanda@gmail.com')</p>
        <input type="text" id="security_answer" name="security_answer" required>
        
        <input type="submit" value="Daftar" class="btn-blue">
    </form>
    
    <hr style="margin-top: 20px;">
    <p>Sudah punya akun? <a href="{url_for('index')}">Login di sini</a></p>
    """
    return render_page("Registrasi", body, error_message=error)

@app.route("/logout")
def logout():
    session.pop('username', None)
    session.pop('role', None)
    return redirect(url_for('index'))

# --- Laman 1: KONSUMEN ---
# (Tidak ada perubahan di bagian ini)

@app.route("/home")
def consumer_home():
    if session.get('role') != 'consumer':
        return redirect(url_for('index'))
    username = session.get('username', 'Tamu')
    body = f"""
    <h2>Selamat Datang, {username}!</h2>
    <p>Kami adalah Perusahaan XYZ, solusi terdepan untuk kebutuhan Anda.</p>
    <h3>Fitur Unggulan Kami:</h3>
    <ul>
        <li>Pelayanan Cepat 24/7</li>
        <li>Produk Berkualitas Tinggi</li>
        <li>Harga Kompetitif</li>
    </ul>
    <hr>
    <h3>Menu Pembelian</h3>
    <form action="/purchase" method="POST">
        <label for="item_name">Nama Barang:</label>
        <input type="text" id="item_name" name="item_name" value="Produk A" required>
        <label for="amount">Jumlah (Rp):</label>
        <input type="number" id="amount" name="amount" value="100000" required>
        <label for="payment_method">Metode Pembayaran:</label>
        <select id="payment_method" name="payment_method">
            <option value="cash">Tunai</option>
            <option value="credit">Kredit</option>
        </select>
        <input type="submit" value="Beli Sekarang">
    </form>
    """
    return render_page("Beranda Konsumen", body)

@app.route("/purchase", methods=['POST'])
def purchase():
    if session.get('role') != 'consumer':
        return redirect(url_for('index'))

    item_name = request.form['item_name']
    amount = float(request.form['amount'])
    payment_method = request.form['payment_method']
    description = f"Penjualan {item_name} kepada {session.get('username')}"

    sales_account = '4101'
    asset_account = '1101' if payment_method == 'cash' else '1102'

    db = get_db()
    with db:
        # --- PERUBAHAN DI SINI ---
        # 1. Buat kode jurnal unik (real-time)
        new_code = generate_journal_code(db)
        
        # 2. Simpan kode baru ke database
        cursor = db.execute("INSERT INTO journal_entries (journal_code, description) VALUES (?, ?)", 
                            (new_code, description))
        entry_id = cursor.lastrowid # Ini adalah 'id' (Nomor)
        # --- PERUBAHAN SELESAI ---
        
        db.execute(
            "INSERT INTO journal_details (entry_id, account_code, debit) VALUES (?, ?, ?)",
            (entry_id, asset_account, amount)
        )
        db.execute(
            "INSERT INTO journal_details (entry_id, account_code, credit) VALUES (?, ?, ?)",
            (entry_id, sales_account, amount)
        )
    return redirect(url_for('consumer_home'))


# --- Laman 2: ADMIN (AKUNTANSI) ---

def get_admin_sidebar_html():
    return """
    <h3>Menu Navigasi</h3>
    
    <p><b>Folder 1: Data & Daftar</b></p>
    <ul>
        <li><a href="/admin/chart-of-accounts">Daftar Akun</a></li>
        <li><a href="/admin/opening-balance">Neraca Saldo Awal</a></li>
        <li><a href="/admin/transactions">Daftar Transaksi</a></li>
    </ul>
    
    <p><b>Folder 2: Jurnal</b></p>
    <ul>
        <li><a href="/admin/general-journal">Jurnal Umum</a></li>
        <li><a href="/admin/inventory-journal">Jurnal Inventory</a></li>
        <li><a href="/admin/adjusting-entries">Jurnal Penyesuaian</a></li>
        <li><a href="/admin/closing-entries">Jurnal Penutup</a></li>
    </ul>
    
    <p><b>Folder 3: Buku Besar</b></p>
    <ul>
        <li><a href="/admin/ledger">Buku Besar</a></li>
        <li><a href="/admin/ledger-ar">Buku Pembantu Piutang</a></li>
        <li><a href="/admin/ledger-ap">Buku Pembantu Utang</a></li>
    </ul>
    
    <p><b>Folder 4: Laporan Keuangan</b></p>
    <ul>
        <li><a href="/admin/income-statement">Laporan Laba Rugi</a></li>
        <li><a href="/admin/balance-sheet">Neraca</a></li>
    </ul>
    """

@app.route("/admin")
def admin_dashboard():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    
    sidebar = get_admin_sidebar_html()
    body = f"""
    <p>Selamat datang di dashboard admin, {session.get('username')}.</p>
    <p>Ini adalah pusat kendali siklus akuntansi.</p>
    <p>Silakan gunakan menu di sebelah kiri untuk navigasi.</p>
    """
    return render_page("Dashboard Admin", body, sidebar_content=sidebar)

# --- Folder 1 Routes (DIPERBARUI) ---

@app.route("/admin/chart-of-accounts")
def chart_of_accounts():
    """Menampilkan daftar akun dengan checkbox dan tombol."""
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
        
    db = get_db()
    accounts = db.execute("SELECT id, account_code, account_name, account_type FROM chart_of_accounts ORDER BY account_code").fetchall()
    
    # Dapatkan pesan error jika ada (dari redirect)
    error = request.args.get('error')
    
    table_rows = ""
    for acc in accounts:
        # Checkbox menggunakan 'account_code' sebagai nilainya
        table_rows += f"""
        <tr>
            <td class="col-pilih"><input type="checkbox" name="selected_codes" value="{acc['account_code']}"></td>
            <td>{acc['account_code']}</td>
            <td>{acc['account_name']}</td>
            <td>{acc['account_type']}</td>
        </tr>
        """
        
    body = f"""
    <a href="/admin/add-account" style="text-decoration: none;">
        <input type="button" value="Input Akun Manual" class="btn-blue">
    </a>
    
    <form action="/admin/delete-accounts" method="POST">
        <table class="table-condensed">
            <thead>
                <tr>
                    <th class="col-pilih">Pilih</th>
                    <th>Kode Akun</th>
                    <th>Nama Akun</th>
                    <th>Tipe Akun</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
        <br>
        <input type="submit" value="Hapus Akun Terpilih" class="btn-red" 
               onclick="return confirm('PERINGATAN: Akun yang sudah digunakan dalam jurnal tidak akan bisa dihapus. Lanjutkan?');">
    </form>
    """
    # Kirim pesan error ke render_page jika ada
    return render_page("Daftar Akun", body, sidebar_content=get_admin_sidebar_html(), error_message=error)

@app.route("/admin/transactions")
def transactions_list():
    """Menampilkan daftar transaksi dengan 'Nomor' dan 'ID Jurnal' baru."""
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
        
    db = get_db()
    
    # Ambil 'id' (untuk Nomor) dan 'journal_code' (untuk ID Jurnal)
    entries = db.execute("SELECT id, journal_code, entry_timestamp, description FROM journal_entries ORDER BY id ASC").fetchall()
    
    table_rows = ""
    for entry in entries:
        entry_date, entry_time = entry['entry_timestamp'].split(' ', 1)
        
        # --- PERUBAHAN DI SINI ---
        table_rows += f"""
        <tr>
            <td class="col-pilih"><input type="checkbox" name="selected_ids" value="{entry['id']}"></td>
            <td class="col-nomor">{entry['id']}</td>           <td>{entry['journal_code']}</td> <td>{entry_date}</td>      
            <td>{entry_time}</td>      
            <td>{entry['description']}</td>
        </tr>
        """
        # --- PERUBAHAN SELESAI ---
        
    body = f"""
    <a href="/admin/add-transaction" style="text-decoration: none;">
        <input type="button" value="Input Transaksi Manual" class="btn-blue">
    </a>
    
    <form action="/admin/delete-transactions" method="POST">
        <table class="table-condensed">
            <thead>
                <tr>
                    <th class="col-pilih">Pilih</th>
                    <th class="col-nomor">Nomor</th>         <th>ID Jurnal</th>     <th>Tanggal</th>
                    <th>Jam</th>
                    <th>Deskripsi</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
        <br>
        <input type="submit" value="Hapus Transaksi Terpilih" class="btn-red"
               onclick="return confirm('Anda yakin ingin menghapus transaksi terpilih? Ini akan menghapus header dan semua detail jurnalnya.');">
    </form>
    """
    return render_page("Daftar Transaksi", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/opening-balance")
def opening_balance():
    """
    Menampilkan halaman statis untuk Neraca Saldo Awal
    dengan format mata uang Rp.
    """
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    
    # Tubuh HTML ini telah dimodifikasi untuk menambahkan 'Rp'
    # dan mengganti nilai kosong dengan 'Rp 0'
    body = f"""
    <style>
        /* Tambahkan style khusus untuk halaman ini */
        .trial-balance-table td {{ 
            padding: 8px; 
            border: 1px solid #ddd;
        }}
        .trial-balance-table .header th {{ 
            background-color: #f2f2f2;
            padding: 8px; 
            border: 1px solid #ddd;
            text-align: left;
        }}
        .account-name {{ 
            padding-left: 20px; 
        }}
        .sub-account {{
            padding-left: 40px;
        }}
        .currency {{ 
            text-align: right; 
            font-family: 'Courier New', Courier, monospace;
            /* Tambahkan padding agar 'Rp' tidak terlalu mepet */
            padding-right: 10px !important; 
        }}
        .description-notes {{
            font-style: italic;
            vertical-align: top;
            width: 30%;
        }}
        .total-row td {{
            font-weight: bold;
            background-color: #f2f2f2;
        }}
    </style>
    
    <table class="trial-balance-table" style="width: 100%; border-collapse: collapse;">
        <thead class="header">
            <tr>
                <th colspan="2">DAFTAR AKUN & NERACA SALDO AWAL</th>
                <th>Debit</th>
                <th>Kredit</th>
                <th>Keterangan</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td colspan="2">Kas</td>
                <td class="currency">Rp 15.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Piutang</td>
                <td class="currency">Rp 3.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Obat obatan</td>
                <td class="currency">Rp 2.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Persediaan bibit ikan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Persediaan pakan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Perlengkapan</td>
                <td class="currency">Rp 2.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Peralatan</td>
                <td class="currency">Rp 5.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2" class="sub-account">Akm. Penyusutan peralatan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 2.500.000,00</td>
                <td class="description-notes"><i>penyusutan 10 thn, garis lurus</i></td>
            </tr>
            <tr>
                <td colspan="2">Kendaraan</td>
                <td class="currency">Rp 50.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2" class="sub-account">Akm. Penyusutan kendaraan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 31.250.000,00</td>
                <td class="description-notes"><i>penyusutan 8 thn, garis lurus</i></td>
            </tr>
            <tr>
                <td colspan="2">Bangunan</td>
                <td class="currency">Rp 50.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2" class="sub-account">Akm. Penyusutan bangunan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 12.500.000,00</td>
                <td class="description-notes"><i>penyusutan 20 thn, garis lurus</i></td>
            </tr>
            <tr>
                <td colspan="2">Tanah</td>
                <td class="currency">Rp 300.000.000,00</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Utang usaha</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 5.000.000,00</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Modal</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 375.750.000,00</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Prive</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Penjualan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Pendapatan lain lain</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">HPP</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Biaya gaji</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Biaya akomodasi</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Biaya listrik dan air</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Biaya penyusutan peralatan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Biaya penyusutan kendaraan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Biaya penyusutan bangunan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Beban perlengkapan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
            <tr>
                <td colspan="2">Biaya obat obatan</td>
                <td class="currency">Rp 0</td>
                <td class="currency">Rp 0</td>
                <td></td>
            </tr>
        </tbody>
        <tfoot>
            <tr class="total-row">
                <td colspan="2" style="text-align: right;">Total</td>
                <td class="currency">Rp 427.000.000,00</td>
                <td class="currency">Rp 427.000.000,00</td>
                <td></td>
            </tr>
        </tfoot>
    </table>
    
    <br>
    
    <div style="font-style: italic;">
        Keterangan:
        Memulai usaha tahun 2020 (sampai 2024 penyusutan berjalan 5 thn)<br>
        Transaksi start pada 1 Jan 2025
    </div>
    """
    
    # Render halaman menggunakan template yang ada
    return render_page("Neraca Saldo Awal", body, sidebar_content=get_admin_sidebar_html())
    
    # Render halaman menggunakan template yang ada
    return render_page("Neraca Saldo Awal", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/add-account", methods=['GET', 'POST'])
def add_account():
    """Formulir untuk menambah akun baru secara manual."""
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        # Proses form
        code = request.form['account_code']
        name = request.form['account_name']
        tipe = request.form['account_type']
        
        try:
            db = get_db()
            with db:
                db.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type) VALUES (?, ?, ?)",
                           (code, name, tipe))
            # Jika sukses, kembali ke daftar akun
            return redirect(url_for('chart_of_accounts'))
        except sqlite3.IntegrityError:
            # Jika gagal (misal: kode duplikat)
            error = f"Gagal menambah akun. Kode Akun '{code}' mungkin sudah ada."
            # Tampilkan form lagi dengan pesan error
            return render_add_account_form(error)
        
    # Jika method GET, tampilkan form
    return render_add_account_form()

def render_add_account_form(error=None):
    """Helper untuk merender form tambah akun."""
    body = f"""
    <form action="/admin/add-account" method="POST">
        <label for="account_code">Kode Akun (Berupa Angka XXXX):</label>
        <input type="text" id="account_code" name="account_code" required>
        
        <label for="account_name">Nama Akun:</label>
        <input type="text" id="account_name" name="account_name" required>
        
        <label for="account_type">Tipe Akun:</label>
        <select id="account_type" name="account_type">
            <option value="Aset Tetap">Aset Tetap</option>
            <option value="Aset Lancar">Aset Lancar</option> 
            <option value="Liabilitas">Liabilitas</option>
            <option value="Ekuitas">Ekuitas</option>
            <option value="Pendapatan">Pendapatan</option>
            <option value="Beban">Beban</option>
        </select>
        
        <input type="submit" value="Simpan Akun Baru">
    </form>
    <a href="{url_for('chart_of_accounts')}">Batal</a>
    """
    return render_page("Input Akun", body, sidebar_content=get_admin_sidebar_html(), error_message=error)

@app.route("/admin/delete-accounts", methods=['POST'])
def delete_accounts():
    """Memproses penghapusan akun yang dipilih."""
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
        
    codes_to_delete = request.form.getlist('selected_codes')
    db = get_db()
    error_msg = None
    
    with db:
        for code in codes_to_delete:
            try:
                # Coba hapus
                db.execute("DELETE FROM chart_of_accounts WHERE account_code = ?", (code,))
            except sqlite3.IntegrityError:
                # Gagal karena ON DELETE RESTRICT (akun sudah dipakai di journal_details)
                error_msg = f"Satu atau lebih akun (termasuk {code}) tidak dapat dihapus karena sudah digunakan dalam jurnal."

    # Redirect kembali ke daftar akun, kirim pesan error jika ada
    if error_msg:
        return redirect(url_for('chart_of_accounts', error=error_msg))
    else:
        return redirect(url_for('chart_of_accounts'))

# --- Rute BARU untuk Input dan Delete Transaksi ---

@app.route("/admin/add-transaction", methods=['GET', 'POST'])
def add_transaction():
    """Formulir untuk menambah transaksi (jurnal) manual."""
    if session.get('role') != 'admin':
        return redirect(url_for('index'))

    db = get_db()
    
    if request.method == 'POST':
        # 1. Ambil data header
        # Ambil tanggal dari form
        entry_date_from_form = request.form['entry_date'] 
        # Tambahkan jam default untuk entri manual
        entry_timestamp_str = f"{entry_date_from_form} 00:00:00"
        
        description = request.form['description']
        
        # ... (sisa kode validasi debit/kredit tidak berubah) ...
        accounts = request.form.getlist('account_code')
        debits = request.form.getlist('debit')
        credits = request.form.getlist('credit')
        
        total_debit = 0.0
        total_credit = 0.0
        details = []
        
        for i in range(len(accounts)):
            if accounts[i]: 
                debit = float(debits[i] or 0)
                credit = float(credits[i] or 0)
                
                if debit > 0 or credit > 0:
                    total_debit += debit
                    total_credit += credit
                    details.append({'account': accounts[i], 'debit': debit, 'credit': credit})

        if total_debit != total_credit or total_debit == 0:
            error = f"Jurnal tidak balance! Total Debit (Rp {total_debit}) harus sama dengan Total Kredit (Rp {total_credit}) dan tidak boleh nol."
            all_accounts = db.execute("SELECT account_code, account_name FROM chart_of_accounts ORDER BY account_code").fetchall()
            return render_add_transaction_form(all_accounts, error)
        
        with db:
            # --- PERUBAHAN DI SINI ---
            # Simpan ke 'entry_timestamp'
            cursor = db.execute("INSERT INTO journal_entries (entry_timestamp, description) VALUES (?, ?)",
                                (entry_timestamp_str, description))
            # --- PERUBAHAN SELESAI ---
            entry_id = cursor.lastrowid
            
            for detail in details:
                db.execute("INSERT INTO journal_details (entry_id, account_code, debit, credit) VALUES (?, ?, ?, ?)",
                           (entry_id, detail['account'], detail['debit'], detail['credit']))

        return redirect(url_for('general_journal'))

    all_accounts = db.execute("SELECT account_code, account_name FROM chart_of_accounts ORDER BY account_code").fetchall()
    return render_add_transaction_form(all_accounts)

def render_add_transaction_form(accounts, error=None):
    """Helper untuk merender form tambah transaksi."""
    
    # Buat <option> untuk dropdown
    account_options = '<option value="">-- Pilih Akun --</option>'
    for acc in accounts:
        account_options += f'<option value="{acc["account_code"]}">{acc["account_code"]} - {acc["account_name"]}</option>'
    
    # Buat 5 baris input statis (karena kita tidak bisa pakai JS untuk 'tambah baris')
    detail_rows = ""
    for i in range(5):
        detail_rows += f"""
        <div style="display: flex; width: 100%; gap: 10px;">
            <select name="account_code" style="flex: 3;">{account_options}</select>
            <input type="number" name="debit" placeholder="Debit" style="flex: 1;" value="0">
            <input type="number" name="credit" placeholder="Kredit" style="flex: 1;" value="0">
        </div>
        """
        
    body = f"""
    <p><b>Catatan:</b> Total Debit dan Kredit harus sama (Balance).</p>
    <form action="/admin/add-transaction" method="POST">
        <label for="entry_date">Tanggal Transaksi:</label>
        <input type="date" id="entry_date" name="entry_date" required>
        
        <label for="description">Deskripsi Transaksi:</label>
        <input type="text" id="description" name="description" required>
        
        <hr>
        <div style="display: flex; width: 100%; gap: 10px;">
            <b style="flex: 3;">Akun</b>
            <b style="flex: 1;">Debit</b>
            <b style="flex: 1;">Kredit</b>
        </div>
        {detail_rows}
        
        <input type="submit" value="Simpan Jurnal">
    </form>
    <a href="{url_for('transactions_list')}">Batal</a>
    """
    return render_page("Input Transaksi Manual", body, sidebar_content=get_admin_sidebar_html(), error_message=error)

@app.route("/admin/delete-transactions", methods=['POST'])
def delete_transactions():
    """Memproses penghapusan transaksi yang dipilih."""
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
        
    ids_to_delete = request.form.getlist('selected_ids')
    db = get_db()
    
    with db:
        for entry_id in ids_to_delete:
            db.execute("DELETE FROM journal_entries WHERE id = ?", (entry_id,))

    return redirect(url_for('transactions_list'))


# --- Folder 2 Routes (Jurnal & Buku Besar) ---
# (Tidak ada perubahan di bagian ini)

@app.route("/admin/general-journal")
def general_journal():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    db = get_db()
    query = """
    SELECT 
        j.entry_timestamp,  -- Ganti dari entry_date
        j.description, 
        c.account_name, 
        d.debit, 
        d.credit
    FROM journal_entries j
    JOIN journal_details d ON j.id = d.entry_id
    JOIN chart_of_accounts c ON d.account_code = c.account_code
    ORDER BY j.id, d.debit DESC;
    """
    details = db.execute(query).fetchall()
    
    table_rows = ""
    for row in details:
        account_cell = row['account_name']
        if row['credit'] > 0:
            account_cell = f"<span style='padding-left: 30px;'>{row['account_name']}</span>"

        # Pisahkan timestamp, tapi hanya gunakan tanggalnya
        entry_date, entry_time = row['entry_timestamp'].split(' ', 1)

        table_rows += f"""
        <tr>
            <td>{entry_date if row['debit'] > 0 else ''}</td>
            
            <td>{row['description'] if row['debit'] > 0 else ''}</td>
            <td>{account_cell}</td>
            <td>{row['debit'] if row['debit'] > 0 else ''}</td>
            <td>{row['credit'] if row['credit'] > 0 else ''}</td>
        </tr>
        """
        
    body = f"""
    <h3>Jurnal Umum</h3>
    <table>
        <thead>
            <tr>
                <th>Tanggal</th>
                <th>Deskripsi</th>
                <th>Akun</th>
                <th>Debit (Rp)</th>
                <th>Kredit (Rp)</th>
            </tr>
        </thead>
        <tbody>
            {table_rows}
        </tbody>
    </table>
    """
    return render_page("Jurnal Umum", body, sidebar_content=get_admin_sidebar_html())

def generate_ledger_html(account_code, db):
    """
    Helper function untuk mengambil data dan merender HTML untuk SATU buku besar.
    Mengembalikan string HTML.
    """
    
    acc_name_row = db.execute("SELECT account_name, account_type FROM chart_of_accounts WHERE account_code = ?", (account_code,)).fetchone()
    if not acc_name_row:
        return f"<p>Error: Akun {account_code} tidak ditemukan.</p>"
    
    acc_name = acc_name_row['account_name']
    acc_type = acc_name_row['account_type']

    # --- PERUBAHAN DI SINI ---
    query = """
    SELECT j.entry_timestamp, j.description, d.debit, d.credit
    FROM journal_details d
    JOIN journal_entries j ON d.entry_id = j.id
    WHERE d.account_code = ?
    ORDER BY j.entry_timestamp; -- Urutkan berdasarkan timestamp
    """
    # --- PERUBAHAN SELESAI ---
    transactions = db.execute(query, (account_code,)).fetchall()
    
    table_rows = ""
    saldo = 0.0
    is_normal_debit = acc_type in ('Aset Lancar', 'Aset Tetap', 'Beban')
    
    for trx in transactions:
        # Pisahkan timestamp, tapi hanya gunakan tanggalnya
        entry_date, entry_time = trx['entry_timestamp'].split(' ', 1)

        if is_normal_debit:
            saldo += trx['debit']
            saldo -= trx['credit']
        else:
            saldo -= trx['debit']
            saldo += trx['credit']
            
        table_rows += f"""
        <tr>
            <td>{entry_date}</td> <td>{trx['description']}</td>
            <td>{trx['debit']}</td>
            <td>{trx['credit']}</td>
            <td>{saldo}</td>
        </tr>
        """

    ledger_html = f"""
    <h3>Buku Besar: {acc_name} (Kode: {account_code})</h3>
    <table>
        <thead>
            <tr>
                <th>Tanggal</th>
                <th>Deskripsi</th>
                <th>Debit (Rp)</th>
                <th>Kredit (Rp)</th>
                <th>Saldo (Rp) - Normal { 'Debit' if is_normal_debit else 'Kredit' }</th>
            </tr>
        </thead>
        <tbody>
            {table_rows}
        </tbody>
        <tfoot>
            <tr>
                <td colspan="4" style="text-align: right;"><b>Saldo Akhir</b></td>
                <td><b>{saldo}</b></td>
            </tr>
        </tfoot>
    </table>
    """
    return ledger_html

def format_currency(value):
    """
    Helper untuk memformat angka menjadi format mata uang Rupiah
    Contoh: 15000.0 -> Rp 15.000,00
    Contoh: -5000.0 -> (Rp 5.000,00)
    """
    # Gunakan f-string formatting untuk pemisah ribuan (,) dan 2 desimal (.)
    # Ganti koma dan titik untuk standar Indonesia
    formatted = "{:,.2f}".format(value).replace(",", "X").replace(".", ",").replace("X", ".")
    
    if value < 0:
        return f'(Rp {formatted.replace("-", "")})'
    else:
        return f'Rp {formatted}'

def get_net_income(db):
    """
    Helper untuk menghitung Laba/Rugi Bersih saat ini.
    Mengembalikan angka (float).
    """
    # 1. Ambil semua total Pendapatan
    pendapatan_query = """
    SELECT SUM(d.credit) - SUM(d.debit) as total
    FROM journal_details d
    JOIN chart_of_accounts c ON d.account_code = c.account_code
    WHERE c.account_type = 'Pendapatan'
    """
    pendapatan_data = db.execute(pendapatan_query).fetchone()
    total_pendapatan = pendapatan_data['total'] if pendapatan_data['total'] else 0.0

    # 2. Ambil semua total Beban
    beban_query = """
    SELECT SUM(d.debit) - SUM(d.credit) as total
    FROM journal_details d
    JOIN chart_of_accounts c ON d.account_code = c.account_code
    WHERE c.account_type = 'Beban'
    """
    beban_data = db.execute(beban_query).fetchone()
    total_beban = beban_data['total'] if beban_data['total'] else 0.0
    
    # 3. Hitung Laba/Rugi
    laba_rugi_bersih = total_pendapatan - total_beban
    return laba_rugi_bersih

@app.route("/admin/ledger")
def ledger():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    
    db = get_db()
    # Dapatkan kode akun dari URL (?code=...)
    account_code = request.args.get('code')
    
    # KONDISI 1: TIDAK ADA KODE AKUN (Tampilkan halaman pilihan)
    if not account_code:
        accounts = db.execute("SELECT account_code, account_name, account_type FROM chart_of_accounts ORDER BY account_type, account_code").fetchall()
        
        # --- PERUBAHAN DIMULAI DI SINI ---
        # Tambahkan Tombol "Tampilkan Semua" yang mengarah ke rute baru
        body = f"""
        <a href="{url_for('ledger_all')}" style="text-decoration: none;">
            <input type="button" value="Tampilkan Semua Buku Besar" class="btn-blue">
        </a>
        <hr>
        <h2>Atau, Pilih Akun Satu per Satu:</h2>
        """
        # --- PERUBAHAN SELESAI ---
        
        current_type = ""
        for acc in accounts:
            # Buat grup berdasarkan Tipe Akun
            if acc['account_type'] != current_type:
                if current_type != "":
                    body += "</ul>" # Tutup list sebelumnya
                body += f"<br><b>{acc['account_type']}</b><ul>" # Mulai list baru
                current_type = acc['account_type']
            
            # Buat tautan ke halaman ini lagi, TAPI dengan parameter code=
            body += f'<li><a href="/admin/ledger?code={acc["account_code"]}">{acc["account_code"]} - {acc["account_name"]}</a></li>'
        
        body += "</ul>" # Tutup list terakhir
        
        return render_page("Pilih Buku Besar", body, sidebar_content=get_admin_sidebar_html())

    # KONDISI 2: ADA KODE AKUN (Tampilkan satu ledger)
    else:
        # Panggil helper function yang baru kita buat
        single_ledger_html = generate_ledger_html(account_code, db)
        
        body = f"""
        <p><a href="/admin/ledger">&larr; Kembali ke Pilihan Akun</a></p>
        {single_ledger_html}
        """
        
        # Ambil nama akun lagi hanya untuk judul halaman
        acc_name_row = db.execute("SELECT account_name FROM chart_of_accounts WHERE account_code = ?", (account_code,)).fetchone()
        acc_name = acc_name_row['account_name'] if acc_name_row else "Error"
        
        return render_page(f"Buku Besar - {acc_name}", body, sidebar_content=get_admin_sidebar_html())
    
@app.route("/admin/ledger-all")
def ledger_all():
    """Rute BARU untuk menampilkan SEMUA buku besar sekaligus."""
    if session.get('role') != 'admin':
        return redirect(url_for('index'))

    db = get_db()
    # Ambil SEMUA akun
    accounts = db.execute("SELECT account_code FROM chart_of_accounts ORDER BY account_code").fetchall()
    
    # Mulai body HTML
    all_ledgers_html = f"""
    <p><a href="{url_for('ledger')}">&larr; Kembali ke Pilihan Akun</a></p>
    <hr>
    """
    
    # Loop melalui setiap akun
    for acc in accounts:
        account_code = acc['account_code']
        # Panggil helper function untuk setiap akun
        all_ledgers_html += generate_ledger_html(account_code, db)
        all_ledgers_html += "<hr>" # Beri pemisah antar buku besar

    # Render halaman penuh
    return render_page("Semua Buku Besar", all_ledgers_html, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/inventory-journal")
def inventory_journal():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    body = """
    <h2>Jurnal Inventory</h2>
    <h3 style="color: #aa0000;">Fitur Tidak Dapat Dibuat</h3>
    <p>Fitur ini tidak dapat dibuat dengan struktur database saat ini.</p>
    <p><b>Alasan:</b> Jurnal inventory (metode perpetual) memerlukan pelacakan <b>kuantitas (Qty)</b> dan 
    <b>harga pokok (Cost)</b> untuk setiap barang (Obat, Pakan, Bibit). Database kita saat ini hanya 
    mencatat nilai Rupiah (Rp) dari transaksi.</p>
    <p><b>Solusi:</b> Ini memerlukan database baru untuk 'Barang' (Items), pelacakan FIFO/Average Cost, 
    dan logika yang sangat kompleks untuk membuat 2 jurnal (Penjualan dan HPP) setiap kali ada penjualan.</p>
    """
    return render_page("Jurnal Inventory", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/adjusting-entries")
def adjusting_entries():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    body = "<h2>Jurnal Penyesuaian</h2><p>Fitur ini sedang dalam pengembangan.</p>"
    return render_page("Jurnal Penyesuaian", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/closing-entries")
def closing_entries():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    body = "<h2>Jurnal Penutup</h2><p>Fitur ini sedang dalam pengembangan.</p>"
    return render_page("Jurnal Penutup", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/ledger-ar")
def ledger_ar():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    body = """
    <h2>Buku Pembantu Piutang</h2>
    <h3 style="color: #aa0000;">Fitur Tidak Dapat Dibuat</h3>
    <p>Fitur ini tidak dapat dibuat dengan struktur database saat ini.</p>
    <p><b>Alasan:</b> Database kita saat ini hanya mencatat transaksi ke akun kontrol 'Piutang Dagang'. 
    Kita tidak menyimpan <b>nama pelanggan</b> yang terkait dengan setiap piutang.</p>
    <p><b>Solusi:</b> Ini memerlukan perubahan database yang besar, termasuk membuat tabel 'Pelanggan' 
    dan menambahkan kolom 'pelanggan_id' di tabel 'journal_details'.</p>
    """
    return render_page("Buku Pembantu Piutang", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/ledger-ap")
def ledger_ap():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    body = """
    <h2>Buku Pembantu Utang</h2>
    <h3 style="color: #aa0000;">Fitur Tidak Dapat Dibuat</h3>
    <p>Fitur ini tidak dapat dibuat dengan struktur database saat ini.</p>
    <p><b>Alasan:</b> Database kita saat ini hanya mencatat transaksi ke akun kontrol 'Utang Usaha'. 
    Kita tidak menyimpan <b>nama supplier</b> yang terkait dengan setiap utang.</p>
    <p><b>Solusi:</b> Ini memerlukan perubahan database yang besar, termasuk membuat tabel 'Supplier' 
    dan menambahkan kolom 'supplier_id' di tabel 'journal_details'.</p>
    """
    return render_page("Buku Pembantu Utang", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/income-statement")
def income_statement():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    
    db = get_db()
    
    # Query untuk mengambil semua akun pendapatan dan beban
    query = """
    SELECT
        c.account_code,
        c.account_name,
        c.account_type,
        SUM(d.debit) as total_debit,
        SUM(d.credit) as total_credit
    FROM journal_details d
    JOIN chart_of_accounts c ON d.account_code = c.account_code
    WHERE c.account_type IN ('Pendapatan', 'Beban')
    GROUP BY c.account_code
    ORDER BY c.account_type DESC, c.account_code;
    """
    accounts = db.execute(query).fetchall()
    
    pendapatan_html = ""
    beban_html = ""
    total_pendapatan = 0.0
    total_beban = 0.0

    for acc in accounts:
        if acc['account_type'] == 'Pendapatan':
            # Saldo normal Pendapatan adalah Kredit
            balance = acc['total_credit'] - acc['total_debit']
            total_pendapatan += balance
            pendapatan_html += f"""
            <tr>
                <td style="padding-left: 20px;">{acc['account_name']}</td>
                <td class="currency">{format_currency(balance)}</td>
            </tr>
            """
        elif acc['account_type'] == 'Beban':
            # Saldo normal Beban adalah Debit
            balance = acc['total_debit'] - acc['total_credit']
            total_beban += balance
            beban_html += f"""
            <tr>
                <td style="padding-left: 20px;">{acc['account_name']}</td>
                <td class="currency">{format_currency(balance)}</td>
            </tr>
            """
            
    # Panggil helper yang kita buat
    laba_rugi_bersih = get_net_income(db)
    
    # Tentukan label hasil
    hasil_label = "Laba Bersih" if laba_rugi_bersih >= 0 else "Rugi Bersih"

    body = f"""
    <style>
        .report-table {{ width: 100%; border-collapse: collapse; }}
        .report-table th, .report-table td {{ padding: 8px; border: 1px solid #ddd; }}
        .report-table th {{ background-color: #f2f2f2; }}
        .header-row {{ font-weight: bold; background-color: #f9f9f9; }}
        .total-row {{ font-weight: bold; background-color: #f2f2f2; border-top: 2px solid #333; }}
        .currency {{ text-align: right; font-family: 'Courier New', Courier; }}
    </style>
    
    <h2>Perusahaan XYZ</h2>
    <h2>Laporan Laba Rugi</h2>
    <p>Untuk Periode yang Berakhir [Tanggal Hari Ini]</p>
    
    <table class="report-table">
        <tr class="header-row">
            <th colspan="2">Pendapatan</th>
        </tr>
        {pendapatan_html}
        <tr class="total-row">
            <td>Total Pendapatan</td>
            <td class="currency">{format_currency(total_pendapatan)}</td>
        </tr>
        
        <tr class="header-row">
            <th colspan="2">Beban</th>
        </tr>
        {beban_html}
        <tr class="total-row">
            <td>Total Beban</td>
            <td class="currency">{format_currency(total_beban)}</td>
        </tr>
        
        <tr class="total-row" style="font-size: 1.1em; background-color: #e0e0e0;">
            <td>{hasil_label}</td>
            <td class="currency">{format_currency(laba_rugi_bersih)}</td>
        </tr>
    </table>
    """
    return render_page("Laporan Laba Rugi", body, sidebar_content=get_admin_sidebar_html())

@app.route("/admin/balance-sheet")
def balance_sheet():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
        
    db = get_db()
    
    # Dapatkan Laba/Rugi Bersih dari helper
    laba_rugi_bersih = get_net_income(db)
    
    # Query untuk semua akun Neraca
    query = """
    SELECT
        c.account_code,
        c.account_name,
        c.account_type,
        SUM(d.debit) as total_debit,
        SUM(d.credit) as total_credit
    FROM journal_details d
    JOIN chart_of_accounts c ON d.account_code = c.account_code
    WHERE c.account_type IN ('Aset Lancar', 'Aset Tetap', 'Liabilitas', 'Ekuitas')
    GROUP BY c.account_code
    ORDER BY c.account_type, c.account_code;
    """
    accounts = db.execute(query).fetchall()

    # Siapkan string HTML untuk setiap bagian
    aset_lancar_html = ""
    aset_tetap_html = ""
    liabilitas_html = ""
    ekuitas_html = ""
    
    total_aset_lancar = 0.0
    total_aset_tetap = 0.0
    total_liabilitas = 0.0
    total_ekuitas_awal = 0.0 # Ekuitas sebelum Laba/Rugi

    for acc in accounts:
        if acc['account_type'] == 'Aset Lancar':
            balance = acc['total_debit'] - acc['total_credit']
            total_aset_lancar += balance
            aset_lancar_html += f'<tr><td>{acc["account_name"]}</td><td class="currency">{format_currency(balance)}</td></tr>'
        
        elif acc['account_type'] == 'Aset Tetap':
            balance = acc['total_debit'] - acc['total_credit']
            total_aset_tetap += balance
            aset_tetap_html += f'<tr><td>{acc["account_name"]}</td><td class="currency">{format_currency(balance)}</td></tr>'

        elif acc['account_type'] == 'Liabilitas':
            balance = acc['total_credit'] - acc['total_debit']
            total_liabilitas += balance
            liabilitas_html += f'<tr><td>{acc["account_name"]}</td><td class="currency">{format_currency(balance)}</td></tr>'
            
        elif acc['account_type'] == 'Ekuitas':
            # Akun Prive (jika ada) akan mengurangi ekuitas
            if 'prive' in acc['account_name'].lower():
                balance = acc['total_debit'] - acc['total_credit']
            else:
                balance = acc['total_credit'] - acc['total_debit']
            total_ekuitas_awal += balance
            ekuitas_html += f'<tr><td>{acc["account_name"]}</td><td class="currency">{format_currency(balance)}</td></tr>'

    # Hitung total
    total_aset = total_aset_lancar + total_aset_tetap
    total_ekuitas_akhir = total_ekuitas_awal + laba_rugi_bersih
    total_liabilitas_ekuitas = total_liabilitas + total_ekuitas_akhir
    
    # Tambahkan Laba/Rugi Bersih ke bagian Ekuitas
    ekuitas_html += f'<tr><td>Laba (Rugi) Bersih</td><td class="currency">{format_currency(laba_rugi_bersih)}</td></tr>'

    body = f"""
    <style>
        .report-table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
        .report-table th, .report-table td {{ padding: 8px; border: 1px solid #ddd; }}
        .report-table th {{ background-color: #f2f2f2; text-align: left; }}
        .header-row {{ font-weight: bold; background-color: #f9f9f9; }}
        .total-row {{ font-weight: bold; background-color: #f2f2f2; border-top: 2px solid #333; }}
        .currency {{ text-align: right; font-family: 'Courier New', Courier; }}
        
        .balance-sheet-container {{ display: flex; width: 100%; gap: 20px; }}
        .bs-left {{ width: 50%; }}
        .bs-right {{ width: 50%; }}
    </style>
    
    <h2 style="text-align: center;">Perusahaan XYZ</h2>
    <h2 style="text-align: center;">Neraca</h2>
    <p style="text-align: center;">Per [Tanggal Hari Ini]</p>
    
    <div class="balance-sheet-container">
        <div class="bs-left">
            <table class="report-table">
                <tr class="header-row"><th colspan="2">ASET</th></tr>
                
                <tr><th colspan="2" style="background-color: #fafafa;">Aset Lancar</th></tr>
                {aset_lancar_html}
                <tr class="total-row"><td>Total Aset Lancar</td><td class="currency">{format_currency(total_aset_lancar)}</td></tr>
                
                <tr><th colspan="2" style="background-color: #fafafa;">Aset Tetap</th></tr>
                {aset_tetap_html}
                <tr class="total-row"><td>Total Aset Tetap</td><td class="currency">{format_currency(total_aset_tetap)}</td></tr>
                
                <tr class="total-row" style="font-size: 1.1em; background-color: #e0e0e0;">
                    <td>TOTAL ASET</td>
                    <td class="currency">{format_currency(total_aset)}</td>
                </tr>
            </table>
        </div>
        
        <div class="bs-right">
            <table class="report-table">
                <tr class="header-row"><th colspan="2">LIABILITAS & EKUITAS</th></tr>
                
                <tr><th colspan="2" style="background-color: #fafafa;">Liabilitas</th></tr>
                {liabilitas_html}
                <tr class="total-row"><td>Total Liabilitas</td><td class="currency">{format_currency(total_liabilitas)}</td></tr>
                
                <tr><th colspan="2" style="background-color: #fafafa;">Ekuitas</th></tr>
                {ekuitas_html}
                <tr class="total-row"><td>Total Ekuitas</td><td class="currency">{format_currency(total_ekuitas_akhir)}</td></tr>
                
                <tr class="total-row" style="font-size: 1.1em; background-color: #e0e0e0;">
                    <td>TOTAL LIABILITAS & EKUITAS</td>
                    <td class="currency">{format_currency(total_liabilitas_ekuitas)}</td>
                </tr>
            </table>
        </div>
    </div>
    """
    return render_page("Neraca", body, sidebar_content=get_admin_sidebar_html())

# --- Menjalankan Aplikasi ---

if __name__ == '__main__':
    # Hapus file DB lama jika ada untuk menerapkan schema baru (ON DELETE)
    import os
    if os.path.exists(DATABASE):
         os.remove(DATABASE)
    
    init_db() 
    app.run(debug=True)
