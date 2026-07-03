from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
import sqlite3
import datetime
import io
import openpyxl

app = Flask(__name__)
app.secret_key = 'a8f3k9s2...random...'

def get_db_connection():
    conn = sqlite3.connect('tennis.db')
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        name = request.form['name']
        phone = request.form['phone']
        gender = request.form['gender']
        court = request.form['court']
        date = request.form['date']
        time_slot = request.form['time_slot']
        
        if len(phone) != 11 or not phone.isdigit():
            flash('❌ خطا: شماره تماس باید دقیقاً ۱۱ رقم باشد!', 'error')
            return redirect(url_for('index'))

        valid_male_courts = ['زمین شماره 1', 'زمین شماره 2', 'زمین شماره 3']
        if gender == 'زن' and court != 'زمین بانوان':
            flash('❌ خطا: بانوان گرامی، لطفاً فقط گزینه "زمین بانوان" را انتخاب کنید.', 'error')
            return redirect(url_for('index'))
        elif gender == 'مرد' and court not in valid_male_courts:
            flash('❌ خطا: آقایان گرامی، لطفاً فقط زمین شماره ۱، ۲ یا ۳ را انتخاب کنید.', 'error')
            return redirect(url_for('index'))

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM bookings WHERE court_name = ? AND booking_date = ? AND time_slot = ? AND is_archived = 0', (court, date, time_slot))
        
        if cursor.fetchone()[0] > 0:
            flash('❌ خطا: این سانس در تاریخ انتخابی قبلاً رزرو شده است!', 'error')
        else:
            cursor.execute('''
                INSERT INTO bookings (customer_name, phone, gender, court_name, booking_date, time_slot)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (name, phone, gender, court, date, time_slot))
            conn.commit()
            flash('✅ رزرو شما با موفقیت ثبت شد!', 'success')
        conn.close()
        return redirect(url_for('index'))
    return render_template('index.html')

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if request.method == 'POST':
        if request.form.get('password') == 'ali.212':
            session['admin_logged_in'] = True
            return redirect(url_for('admin'))
        else:
            flash('❌ رمز عبور اشتباه است!', 'error')
            return render_template('admin.html', logged_in=False)
    
    if not session.get('admin_logged_in'):
        return render_template('admin.html', logged_in=False)

    conn = get_db_connection()
    # دریافت رزروهای فعال (کمتر از 24 ساعت)
    active_bookings = conn.execute('''
        SELECT * FROM bookings 
        WHERE is_archived = 0 
        ORDER BY booking_date DESC, time_slot ASC
    ''').fetchall()
    
    # دریافت رزروهای بایگانی شده
    archived_bookings = conn.execute('''
        SELECT * FROM bookings 
        WHERE is_archived = 1 
        ORDER BY created_at DESC
    ''').fetchall()
    conn.close()
    
    return render_template('admin.html', logged_in=True, active=active_bookings, archived=archived_bookings)

@app.route('/admin/archive')
def archive_old():
    if not session.get('admin_logged_in'):
        return redirect(url_for('admin'))
    
    conn = get_db_connection()
    # محاسبه زمان 24 ساعت پیش
    limit_time = (datetime.datetime.utcnow() - datetime.timedelta(hours=24)).strftime('%Y-%m-%d %H:%M:%S')
    conn.execute('UPDATE bookings SET is_archived = 1 WHERE created_at < ? AND is_archived = 0', (limit_time,))
    conn.commit()
    conn.close()
    flash('✅ رزروهای قدیمی‌تر از ۲۴ ساعت با موفقیت بایگانی شدند.', 'success')
    return redirect(url_for('admin'))

@app.route('/admin/delete/<int:id>', methods=['POST'])
def delete_booking(id):
    if not session.get('admin_logged_in'):
        return redirect(url_for('admin'))
    
    conn = get_db_connection()
    conn.execute('DELETE FROM bookings WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    flash('✅ رزرو با موفقیت حذف شد.', 'success')
    return redirect(url_for('admin'))

@app.route('/admin/edit/<int:id>', methods=['GET', 'POST'])
def edit_booking(id):
    if not session.get('admin_logged_in'):
        return redirect(url_for('admin'))
    
    conn = get_db_connection()
    booking = conn.execute('SELECT * FROM bookings WHERE id = ?', (id,)).fetchone()
    
    if request.method == 'POST':
        conn.execute('''
            UPDATE bookings 
            SET customer_name = ?, phone = ?, gender = ?, court_name = ?, booking_date = ?, time_slot = ?
            WHERE id = ?
        ''', (
            request.form['name'], request.form['phone'], request.form['gender'],
            request.form['court'], request.form['date'], request.form['time_slot'], id
        ))
        conn.commit()
        conn.close()
        flash('✅ اطلاعات رزرو با موفقیت ویرایش شد.', 'success')
        return redirect(url_for('admin'))
    
    conn.close()
    return render_template('edit.html', booking=booking)

@app.route('/admin/export')
def export_excel():
    if not session.get('admin_logged_in'):
        return redirect(url_for('admin'))
    
    conn = get_db_connection()
    # خروجی اکسل از رزروهای بایگانی شده
    bookings = conn.execute('SELECT * FROM bookings WHERE is_archived = 1 ORDER BY created_at DESC').fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "بایگانی رزروها"
    
    # هدرهای فارسی
    headers = ["شماره", "نام ورزشکار", "جنسیت", "تلفن", "زمین", "تاریخ رزرو", "سانس", "زمان ثبت در سیستم"]
    ws.append(headers)
    
    for b in bookings:
        ws.append([b['id'], b['customer_name'], b['gender'], b['phone'], b['court_name'], b['booking_date'], b['time_slot'], b['created_at']])

    # ذخیره در حافظه و ارسال به کاربر
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='archive_bookings.xlsx'
    )

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged_in', None)
    return redirect(url_for('admin'))

if __name__ == '__main__':
    app.run(debug=False)