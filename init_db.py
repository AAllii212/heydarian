import sqlite3

def init_db():
    conn = sqlite3.connect('tennis.db')
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            gender TEXT NOT NULL,
            court_name TEXT NOT NULL,
            booking_date TEXT NOT NULL,
            time_slot TEXT NOT NULL,
            is_archived INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    conn.commit()
    conn.close()
    print("✅ دیتابیس با موفقیت ساخته/به‌روزرسانی شد!")

if __name__ == '__main__':
    init_db()